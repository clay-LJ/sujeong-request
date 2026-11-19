import '@fontsource/noto-sans-kr/400.css';
import './style.css';
import { parseZip, buildPresentation, renderSlide, RECOMMENDED_ZIP_LIMITS } from '@aiden0z/pptx-renderer';
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { toCanvas, getFontEmbedCSS } from 'html-to-image';
import { jsPDF } from 'jspdf';
pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;
const $ = s => document.querySelector(s);
$('#app').innerHTML = `
<header class="top"><div class="brand"><span class="logo">수</span><div><h1>수정요청서</h1><div class="tagline">위치를 찍고, 요청을 명확하게.</div></div><span class="badge">기기 안에서 처리</span></div><div class="actions"><button id="replace" class="hidden">새 파일</button><input id="title" class="title-input hidden" aria-label="수정요청서 제목" value="수정요청서" maxlength="80"><button id="export" class="primary hidden">PDF 내보내기</button></div></header>
<div id="status" role="status" aria-live="polite"></div>
<main id="welcome"><div class="intro"><div class="eyebrow">REVIEW → REQUEST</div><h2>수정할 곳을 콕, 전달은 한 번에.</h2><p>이미지, PDF, PowerPoint를 열고 번호별 요청을 적어 보세요.<br>원본 페이지와 요청 내용이 담긴 PDF로 바로 전달할 수 있어요.</p></div><div class="drop" id="drop"><div class="drop-icon">▧ +</div><label class="upload" for="file">파일 선택하기</label><p class="small">또는 여기에 파일을 놓아 주세요 · 최대 40MB / 60페이지</p><div class="formats"><span>이미지</span><span>PDF</span><span>PPTX 직접 열기</span></div><div class="samples">합성 예제로 시험하기 <a href="./samples/review-demo.pptx" download>PPTX</a><a href="./samples/review-demo.pdf" download>PDF</a><a href="./samples/review-demo.png" download>이미지</a></div></div><div class="steps"><span><b>01</b>파일 열기</span><span><b>02</b>번호 + 요청 작성</span><span><b>03</b>PDF 내보내기</span></div></main>
<input id="file" type="file" class="hidden" accept=".pptx,.pdf,.png,.jpg,.jpeg,.webp,.gif,.ppt">
<section id="workspace" class="workspace hidden"><div class="viewer"><div class="toolbar"><span class="file" id="filename"></span><div class="nav"><button id="prev" aria-label="이전 페이지">←</button><span id="page-label"></span><button id="next" aria-label="다음 페이지">→</button><select id="zoom" aria-label="확대 비율"><option value="1">화면에 맞춤</option><option value="1.5">150%</option><option value="2">200%</option></select></div></div><div class="stage-scroll" id="scroll"><div class="stage" id="stage" tabindex="0" aria-label="원본 페이지. 클릭하여 수정 번호 추가"><img id="original" alt="원본 페이지 미리보기" draggable="false"><div id="markers"></div></div></div><div class="hint">원본을 클릭하면 번호가 추가됩니다. 번호를 선택해 내용을 편집하거나 삭제하세요.</div></div><aside class="requests"><div class="request-head"><h3>수정 요청</h3><span class="count" id="count">0개</span></div><div id="requests" class="request-list"></div><div class="request-foot">모든 페이지의 요청이 표시됩니다. 입력 내용은 현재 탭 메모리에만 남습니다. 새로고침 전 PDF를 저장하세요.</div></aside></section>
<details class="limits"><summary>지원 범위 · 정확도 · 개인정보</summary><p>파일은 브라우저 메모리에서만 처리됩니다. 원본 업로드, 서버 저장, 분석 수집 기능이 없습니다. 새로고침하거나 탭을 닫으면 작업은 사라집니다. 사이트 자체는 GitHub Pages에서 제공됩니다.</p><p>PPTX는 직접 렌더링하지만 PowerPoint와 완전히 동일하지 않습니다. 글꼴 대체, 줄바꿈, 복잡한 차트·SmartArt·효과의 모양이 달라질 수 있습니다. 애니메이션·전환, 3D, 실행형 OLE, 슬라이드 노트와 일부 수식·EMF/WMF는 지원하지 않습니다. 외부 연결 미디어는 차단합니다. 암호 파일과 구형 .ppt는 지원하지 않습니다.</p><p>이미지는 PNG·JPEG·WebP·GIF(첫 프레임)를 지원합니다. PDF는 보이는 페이지를 읽으며 양식·첨부·주석의 편집은 지원하지 않습니다. 내보낸 PDF는 한국어 외형 보존을 위한 이미지 기반 문서이므로 텍스트 검색·선택이 되지 않습니다. 원본을 편집하거나 원래 PPTX로 저장하지 않습니다. 내보내기 전에 미리보기의 누락과 위치를 확인해 주세요.</p></details><footer>수정요청서 · 가볍게 검토하고, 분명하게 전달하기</footer><div id="render-host" class="render-host"></div>`;
let pages = [], requests = [], current = 0, nextId = 1, active = null, filename = '', busy = false, generation = 0, pdfDoc = null, pdfTask = null, presentation = null, embeddedCSS;
const status = (s, error=false) => { $('#status').textContent=s; $('#status').classList.toggle('error',error); };
function lock(value){busy=value; for(const id of ['export','replace','prev','next','zoom','file']) $('#'+id).disabled=value; $('#stage').style.pointerEvents=value?'none':''; $('#requests').inert=value;}
function message(error){return error instanceof Error ? error.message : String(error);}
async function load(file){
 if(!file || busy)return;
 const ext=file.name.split('.').pop().toLowerCase();
 if(ext==='ppt'){status('구형 .ppt 파일은 지원하지 않습니다. PowerPoint 2007 이후의 .pptx 파일을 선택해 주세요.',true);return;}
 if(!['pptx','pdf','png','jpg','jpeg','webp','gif'].includes(ext)){status('PPTX, PDF, PNG, JPEG, WebP, GIF 파일을 선택해 주세요.',true);return;}
 if(file.size>40*1024*1024){status('파일 크기는 최대 40MB입니다.',true);return;}
 if(requests.length && !window.confirm('새 파일을 열면 현재 요청 내용이 지워집니다. 계속할까요?'))return;
 lock(true);status('파일을 기기 안에서 읽고 있습니다…');
 const oldTask=pdfTask; let candidateTask=null;
 try{
  const buffer=await file.arrayBuffer(); let candidate, newPdf=null, newPres=null;
  if(ext==='pptx'){
   const files=await parseZip(buffer,RECOMMENDED_ZIP_LIMITS);newPres=buildPresentation(files);
   candidate=newPres.slides.map(()=>({width:newPres.width,height:newPres.height,kind:'pptx'}));
  }else if(ext==='pdf'){
   candidateTask=pdfjs.getDocument({data:new Uint8Array(buffer),isEvalSupported:false,useSystemFonts:true,cMapUrl:new URL('cmaps/',document.baseURI).href,cMapPacked:true,standardFontDataUrl:new URL('standard_fonts/',document.baseURI).href,wasmUrl:new URL('wasm/',document.baseURI).href});newPdf=await candidateTask.promise;
   candidate=Array.from({length:newPdf.numPages},()=>({kind:'pdf'}));
  }else{
   const src=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(new Error('이미지를 읽을 수 없습니다.'));r.readAsDataURL(file);});
   const img=await image(src);if(img.width*img.height>50000000)throw new Error('이미지 해상도가 너무 큽니다. 5천만 픽셀 이하 이미지를 선택해 주세요.');
   const canvas=document.createElement('canvas');const scale=Math.min(1,1800/Math.max(img.width,img.height));canvas.width=img.width*scale;canvas.height=img.height*scale;canvas.getContext('2d').drawImage(img,0,0,canvas.width,canvas.height);
   candidate=[{kind:'image',width:canvas.width,height:canvas.height,src:canvas.toDataURL('image/png')}];
  }
  if(!candidate.length || candidate.length>60){throw new Error('1~60페이지 파일을 지원합니다.');}
  pages=candidate;pdfDoc=newPdf;pdfTask=candidateTask;presentation=newPres;requests=[];nextId=1;current=0;active=null;filename=file.name;generation++;
  await oldTask?.destroy();$('#welcome').classList.add('hidden');$('#workspace').classList.remove('hidden');for(const id of ['replace','title','export'])$('#'+id).classList.remove('hidden');$('#filename').textContent=filename;
  await showPage(); renderRequests();status(pages[current].warning||(ext==='pptx'?'PPTX를 직접 열었습니다. 글꼴·도형의 정확도와 누락 여부를 미리보기에서 확인해 주세요.':'파일을 열었습니다. 수정할 위치를 클릭해 주세요.'),!!pages[current].warning);
 }catch(e){if(candidateTask && candidateTask!==pdfTask)await candidateTask.destroy();status('파일을 열지 못했습니다: '+message(e),true);}
 finally{lock(false);updateNav();$('#file').value='';}
}
function image(src){return new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>reject(new Error('이미지 표시 실패'));img.src=src;});}
async function pageImage(index){
 const p=pages[index];if(p.src)return p.src;
 if(p.kind==='pdf'){
  const page=await pdfDoc.getPage(index+1);const v=page.getViewport({scale:1});const viewport=page.getViewport({scale:Math.min(2,1800/Math.max(v.width,v.height))});
  const c=document.createElement('canvas');c.width=Math.ceil(viewport.width);c.height=Math.ceil(viewport.height);await page.render({canvasContext:c.getContext('2d'),viewport}).promise;p.width=c.width;p.height=c.height;p.src=c.toDataURL('image/png');page.cleanup();
 }else{
  let errors=[];const h=renderSlide(presentation,presentation.slides[index],{externalMedia:'block',pdfjs:false,onNodeError:(id,e)=>errors.push(id+': '+message(e))});
  const host=$('#render-host');host.replaceChildren(h.element);
  try{
   await h.ready;await document.fonts.ready;
   // Use bundled Korean fallback even when the source names an unavailable Office font.
   for(const el of [h.element,...h.element.querySelectorAll('*')]){if(el.style.fontFamily)el.style.fontFamily += ', "Noto Sans KR", sans-serif';}
   await document.fonts.load('16px "Noto Sans KR"',h.element.textContent);await document.fonts.ready;
   embeddedCSS ??= await getFontEmbedCSS(h.element,{preferredFontFormat:'woff2'});
   const canvas=await toCanvas(h.element,{pixelRatio:Math.min(2,1800/Math.max(p.width,p.height)),fontEmbedCSS:embeddedCSS,backgroundColor:'#ffffff',cacheBust:false});
   p.src=canvas.toDataURL('image/png');p.warning=errors.length?'일부 PPTX 요소를 표시하지 못했습니다: '+errors.join(' / ').slice(0,300):'';
  }finally{h.dispose();host.replaceChildren();}
 }
 return p.src;
}
async function showPage(){const token=generation;const index=current;$('#page-label').textContent=`${index+1} / ${pages.length}`;$('#stage').style.visibility='hidden';const src=await pageImage(index);if(token!==generation||index!==current)return;$('#original').src=src;await $('#original').decode();resize();renderMarkers();$('#stage').style.visibility='visible';updateNav();}
function resize(){if(!pages.length)return;const p=pages[current];const available=$('#scroll').clientWidth-(innerWidth<850?36:56);$('#stage').style.width=Math.max(80,available)*Number($('#zoom').value)+'px';$('#stage').style.aspectRatio=`${p.width} / ${p.height}`;}
new ResizeObserver(resize).observe($('#scroll'));
function updateNav(){$('#prev').disabled=busy||current===0;$('#next').disabled=busy||current>=pages.length-1;}
async function navigate(index){if(busy||index<0||index>=pages.length)return;current=index;lock(true);try{await showPage();if(pages[current].warning)status(pages[current].warning,true);}catch(e){status('페이지 표시 실패: '+message(e),true);}finally{lock(false);updateNav();}}
function renderMarkers(){$('#markers').replaceChildren();for(const r of requests.filter(r=>r.page===current)){const b=document.createElement('button');b.className='marker'+(r.id===active?' selected':'');b.style.left=r.x*100+'%';b.style.top=r.y*100+'%';b.textContent=r.id;b.setAttribute('aria-label',`수정 요청 ${r.id} 편집`);b.onclick=e=>{e.stopPropagation();active=r.id;renderRequests();renderMarkers();$('#request-'+r.id)?.focus();};$('#markers').append(b);}}
function renderRequests(){
 $('#count').textContent=requests.length+'개';const list=$('#requests');list.replaceChildren();if(!requests.length){const e=document.createElement('div');e.className='empty';e.textContent='아직 수정 요청이 없습니다.\n왼쪽 원본을 클릭해 첫 번호를 추가하세요.';list.append(e);return;}
 for(const r of requests){const card=document.createElement('div');card.className='card'+(r.id===active?' active':'');const top=document.createElement('div');top.className='card-top';const label=document.createElement('span');const num=document.createElement('b');num.className='num';num.textContent=r.id;const link=document.createElement('button');link.className='page-link';link.textContent=`${r.page+1}페이지로 이동`;link.onclick=()=>navigate(r.page);label.append(num,link);const del=document.createElement('button');del.className='delete';del.textContent='삭제';del.setAttribute('aria-label',`수정 요청 ${r.id} 삭제`);del.onclick=()=>{requests=requests.filter(x=>x.id!==r.id);if(active===r.id)active=null;renderRequests();renderMarkers();};top.append(label,del);const input=document.createElement('textarea');input.id='request-'+r.id;input.maxLength=5000;input.placeholder='예: 제목을 “2026년 운영 계획”으로 변경해 주세요.';input.setAttribute('aria-label',`수정 요청 ${r.id} 내용`);input.value=r.text;input.oninput=()=>r.text=input.value;input.onfocus=()=>{active=r.id;renderMarkers();for(const c of list.children)c.classList.remove('active');card.classList.add('active');};card.append(top,input);list.append(card);}
}
$('#stage').onclick=e=>{if(busy||!pages.length)return;const rect=$('#stage').getBoundingClientRect();const r={id:nextId++,page:current,x:Math.max(0,Math.min(1,(e.clientX-rect.left)/rect.width)),y:Math.max(0,Math.min(1,(e.clientY-rect.top)/rect.height)),text:''};requests.push(r);active=r.id;renderMarkers();renderRequests();$('#request-'+r.id).focus();};
$('#prev').onclick=()=>navigate(current-1);$('#next').onclick=()=>navigate(current+1);$('#zoom').onchange=resize;$('#file').onchange=e=>load(e.target.files[0]);$('#replace').onclick=()=>$('#file').click();
const drop=$('#drop');drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag');};drop.ondragleave=()=>drop.classList.remove('drag');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('drag');if(e.dataTransfer.files.length!==1){status('한 번에 파일 하나를 선택해 주세요.',true);return;}load(e.dataTransfer.files[0]);};
window.addEventListener('beforeunload',e=>{if(requests.length){e.preventDefault();e.returnValue='';}});
function lines(ctx,text,width){const result=[];for(const para of text.split('\n')){let line='';for(const ch of para){if(ctx.measureText(line+ch).width>width && line){result.push(line);line='';}line+=ch;}result.push(line);}return result;}
async function exportPDF(){if(busy||!pages.length)return;lock(true);status('수정요청서 PDF를 만들고 있습니다…');
 try{
  await document.fonts.load('28px "Noto Sans KR"',requests.map(r=>r.text).join('')+filename+$('#title').value);await document.fonts.ready;
  const pdf=new jsPDF({unit:'mm',format:'a4',compress:true});let outputCount=0;
  const c=document.createElement('canvas');c.width=1240;c.height=1754;const ctx=c.getContext('2d');let y=0;
  const text=(str,x,yy,size=24,color='#25334b')=>{ctx.font=`${size}px "Noto Sans KR"`;ctx.fillStyle=color;ctx.fillText(str,x,yy);};
  function newSheet(page,continuation=false){ctx.fillStyle='white';ctx.fillRect(0,0,c.width,c.height);text($('#title').value||'수정요청서',64,82,36);text(`${page+1} / ${pages.length}페이지${continuation?' · 요청 내용 계속':''}`,64,125,20,'#667085');ctx.font='18px "Noto Sans KR"';let name=filename;while(ctx.measureText(name).width>750)name=name.slice(0,-2)+'…';text(name,360,125,18,'#667085');ctx.strokeStyle='#dfe5ee';ctx.beginPath();ctx.moveTo(64,150);ctx.lineTo(1176,150);ctx.stroke();y=190;}
  function flush(){text('브라우저 미리보기 기준 · 이미지 기반 PDF',64,1700,16,'#8590a2');text(String(++outputCount),1140,1700,18,'#8590a2');if(outputCount>1)pdf.addPage();pdf.addImage(c.toDataURL('image/jpeg',0.94),'JPEG',0,0,210,297);}
  for(let i=0;i<pages.length;i++){
   status(`PDF 작성 중 · ${i+1} / ${pages.length}페이지`);newSheet(i);const img=await image(await pageImage(i));const scale=Math.min(1112/img.width,800/img.height);const w=img.width*scale,h=img.height*scale,x=(1240-w)/2,top=y;ctx.drawImage(img,x,top,w,h);ctx.strokeStyle='#d6dce6';ctx.strokeRect(x,top,w,h);
   const rs=requests.filter(r=>r.page===i);for(const r of rs){const mx=x+r.x*w,my=top+r.y*h;ctx.beginPath();ctx.arc(mx,my,19,0,Math.PI*2);ctx.fillStyle='#e44935';ctx.fill();ctx.strokeStyle='white';ctx.lineWidth=3;ctx.stroke();ctx.textAlign='center';text(String(r.id),mx,my+7,21,'white');ctx.textAlign='left';}y+=h+50;text('수정 요청',64,y,26);y+=45;
   if(!rs.length){text('이 페이지에 등록된 수정 요청이 없습니다.',64,y,22,'#667085');}
   for(const r of rs){ctx.font='24px "Noto Sans KR"';const wrapped=lines(ctx,`${r.id}. ${r.text.trim()||'(요청 내용 미입력)'}`,1080);for(let n=0;n<wrapped.length;n++){if(y>1620){flush();newSheet(i,true);}text(wrapped[n],70,y,24);y+=38;}y+=22;}
   flush();
  }
  pdf.save('수정요청서.pdf');status(`PDF를 저장했습니다. 원본 ${pages.length}페이지 · 요청 ${requests.length}개 · 출력 ${outputCount}페이지`);
 }catch(e){status('PDF 내보내기 실패: '+message(e),true);}finally{lock(false);updateNav();}
}
$('#export').onclick=exportPDF;
