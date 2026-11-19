import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
const samples='tests/fixtures/';
async function ready(page){await expect(page.locator('#export')).toBeEnabled({timeout:60000});await expect(page.locator('#stage')).toHaveCSS('visibility','visible');await expect(page.locator('#original')).toHaveJSProperty('complete',true);}
async function clickPoint(page,x=.3,y=.55){const b=await page.locator('#stage').boundingBox();await page.mouse.click(b.x+b.width*x,b.y+b.height*y);}
for(const [ext,count] of [['pptx',3],['pdf',2],['png',1]])test(`${ext}: import, navigation, comments, coordinates, PDF, privacy`,async({page},info)=>{
 const seen=[],errors=[];page.on('request',r=>seen.push({url:r.url(),method:r.method(),bytes:r.postDataBuffer()?.length||0}));page.on('pageerror',e=>errors.push(e.message));
 await page.goto('./');await page.locator('#file').setInputFiles(samples+'review-demo.'+ext);await ready(page);await expect(page.locator('#page-label')).toHaveText(`1 / ${count}`);
 await clickPoint(page);await page.getByRole('textbox',{name:'수정 요청 1 내용',exact:true}).fill('제목을 “새 운영 계획”으로 변경해 주세요. 한글 보존 확인.');
 await page.locator('#zoom').selectOption('2');await page.setViewportSize({width:1100,height:900});const ratio=await page.locator('.marker').evaluate(e=>{const m=e.getBoundingClientRect(),p=e.parentElement.parentElement.getBoundingClientRect();return{x:(m.x+m.width/2-p.x)/p.width,y:(m.y+m.height/2-p.y)/p.height};});expect(ratio.x).toBeCloseTo(.3,2);expect(ratio.y).toBeCloseTo(.55,2);
 await page.locator('#zoom').selectOption('1');
 if(count>1){await page.locator('#next').click();await ready(page);await expect(page.locator('#page-label')).toHaveText(`2 / ${count}`);await expect(page.locator('.marker')).toHaveCount(0);await clickPoint(page,.7,.4);await page.getByRole('textbox',{name:'수정 요청 2 내용',exact:true}).fill('표의 담당 항목을 확인해 주세요.');await page.locator('#prev').click();await ready(page);await expect(page.locator('.marker')).toHaveCount(1);}
 await page.getByRole('textbox',{name:'수정 요청 1 내용',exact:true}).fill('제목을 “새 운영 계획”으로 변경해 주세요. 한글 보존 확인. 수정 완료.');
 const downloadPromise=page.waitForEvent('download');await page.locator('#export').click();const download=await downloadPromise;const path=info.outputPath(ext+'-request.pdf');await download.saveAs(path);expect((await fs.stat(path)).size).toBeGreaterThan(20000);await expect(page.locator('#status')).toContainText('PDF를 저장했습니다');
 await page.screenshot({path:info.outputPath(ext+'-desktop.png'),fullPage:true});await page.setViewportSize({width:390,height:844});expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBeTruthy();await page.screenshot({path:info.outputPath(ext+'-mobile.png'),fullPage:true});
 expect(errors).toEqual([]);const bad=seen.filter(r=>!r.url.startsWith(new URL(page.url()).origin)&&!r.url.startsWith('blob:')&&!r.url.startsWith('data:'));expect(bad).toEqual([]);expect(seen.filter(r=>!['GET'].includes(r.method))).toEqual([]);await fs.writeFile(info.outputPath('network.json'),JSON.stringify(seen,null,2));
 if(process.env.BASE_URL)return;const bytes=Array.from(await fs.readFile(path));const rendered=await page.evaluate(async b=>{const m=await import('/tests/pdf-inspect.js');return m.inspect(b);},bytes);expect(rendered.length).toBe(count);for(let i=0;i<rendered.length;i++)await fs.writeFile(info.outputPath(`pdf-page-${i+1}.png`),Buffer.from(rendered[i],'base64'));
});
test('legacy and malformed input show clear errors',async({page})=>{await page.goto('./');await page.locator('#file').setInputFiles({name:'legacy.ppt',mimeType:'application/vnd.ms-powerpoint',buffer:Buffer.from('legacy')});await expect(page.locator('#status')).toContainText('구형 .ppt');await page.locator('#file').setInputFiles({name:'bad.pptx',mimeType:'application/octet-stream',buffer:Buffer.from('bad')});await expect(page.locator('#status')).toContainText('파일을 열지 못했습니다');});
test('long Korean request paginates without losing content',async({page},info)=>{await page.goto('./');await page.locator('#file').setInputFiles(samples+'review-demo.png');await ready(page);await clickPoint(page);await page.locator('#requests textarea').fill('긴 한국어 수정 요청입니다. 여러 줄로 나누어 PDF 출력 누락을 확인합니다.\n'.repeat(70));const dp=page.waitForEvent('download');await page.locator('#export').click();const d=await dp;await d.saveAs(info.outputPath('long-request.pdf'));await expect(page.locator('#status')).toContainText('PDF를 저장했습니다');const label=await page.locator('#status').textContent();expect(Number(label.match(/출력 (\d+)/)[1])).toBeGreaterThan(1);});
test('replace PDF, delete marker and keep numbering unique',async({page})=>{await page.goto('./');await page.locator('#file').setInputFiles(samples+'review-demo.pdf');await ready(page);await clickPoint(page);await page.locator('#requests textarea').fill('삭제할 요청');await page.getByRole('button',{name:'수정 요청 1 삭제',exact:true}).click();await expect(page.locator('.marker')).toHaveCount(0);await clickPoint(page);await expect(page.locator('.marker')).toHaveText('2');page.on('dialog',d=>d.accept());await page.locator('#file').setInputFiles(samples+'review-demo.png');await ready(page);await expect(page.locator('#page-label')).toHaveText('1 / 1');await expect(page.locator('.marker')).toHaveCount(0);await expect(page.locator('#status')).not.toContainText('실패');});
test('PPTX linked media never requests external URL',async({page})=>{const {default:JSZip}=await import('jszip');const zip=await JSZip.loadAsync(await fs.readFile(samples+'review-demo.pptx'));const xml='ppt/slides/slide1.xml',rel='ppt/slides/_rels/slide1.xml.rels';zip.file(xml,(await zip.file(xml).async('string')).replace('</p:spTree>','<p:pic><p:nvPicPr><p:cNvPr id="999" name="external"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:link="rId999"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>'));zip.file(rel,(await zip.file(rel).async('string')).replace('</Relationships>','<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/privacy-probe.png" TargetMode="External"/></Relationships>'));const requests=[];page.on('request',r=>requests.push(r.url()));await page.goto('./');await page.locator('#file').setInputFiles({name:'synthetic-external.pptx',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',buffer:await zip.generateAsync({type:'nodebuffer'})});await ready(page);expect(requests.filter(u=>u.includes('example.invalid'))).toEqual([]);});

test('optional cover includes local logo and Korean handoff details', async ({page}, info) => {
  const traffic = [];
  page.on('request', request => traffic.push(request));
  await page.setViewportSize({width:390,height:844});
  await page.goto('./');
  await expect(page.getByText('합성 예제로 시험하기')).toHaveCount(0);
  const oldSample = new URL('./samples/review-demo.png', page.url());
  const response = await page.request.get(oldSample.href);
  expect(response.headers()['content-type']).not.toContain('image/png');
  await page.locator('#file').setInputFiles(samples+'review-demo.png');
  await ready(page);
  await page.locator('#title').fill('브랜드 소개서 수정요청서');
  await page.locator('#cover-open').click();
  await page.getByLabel('PDF에 표지 추가').check();
  await page.getByLabel('프로젝트', {exact:true}).fill('브랜드 운영 자료 개편');
  await page.getByLabel('수신', {exact:true}).fill('디자인 제작팀');
  await page.getByLabel('발신', {exact:true}).fill('브랜드 운영팀');
  await page.getByLabel('표지 문구').fill('표시한 수정 사항을 검토해 주세요.\n반영본은 작업 담당자에게 전달해 주세요.');
  await page.locator('#cover-file').setInputFiles(samples+'review-demo.png');
  await expect(page.locator('#cover-logo')).toBeVisible();
  await page.screenshot({path:info.outputPath('cover-dialog.png'),fullPage:true});
  await page.getByRole('button',{name:'편집으로 돌아가기'}).click();
  await clickPoint(page);
  await page.getByRole('textbox',{name:'수정 요청 1 내용',exact:true}).fill('첫 장 제목을 최신 문구로 바꿔 주세요.');
  const saved = page.waitForEvent('download');
  await page.locator('#export').click();
  const download = await saved;
  expect(download.suggestedFilename()).toBe('브랜드 소개서 수정요청서.pdf');
  await expect(page.locator('#status')).toContainText('출력 2페이지');
  const path = info.outputPath('cover-report.pdf');
  await download.saveAs(path);
  expect(traffic.filter(request => request.method() !== 'GET' ||
    ![new URL(page.url()).origin,'blob:','data:'].some(prefix => request.url().startsWith(prefix)))).toEqual([]);
  if (!process.env.BASE_URL) {
    const bytes = Array.from(await fs.readFile(path));
    const rendered = await page.evaluate(async data => {
      const module = await import('/tests/pdf-inspect.js');
      return module.inspect(data);
    },bytes);
    expect(rendered).toHaveLength(2);
    for (const [index,png] of rendered.entries()) {
      await fs.writeFile(info.outputPath('report-'+(index+1)+'.png'),Buffer.from(png,'base64'));
    }
  }
  await page.locator('#cover-open').click();
  await page.getByLabel('PDF에 표지 추가').uncheck();
  await page.getByRole('button',{name:'편집으로 돌아가기'}).click();
  const withoutCover = page.waitForEvent('download');
  await page.locator('#export').click();
  await withoutCover;
  await expect(page.locator('#status')).toContainText('출력 1페이지');
});
