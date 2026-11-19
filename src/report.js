import { jsPDF } from 'jspdf';

const navy = '#172d52';
const blue = '#2857a2';
const muted = '#52647e';
const border = '#dce4ef';
const left = 72;
const right = 1168;
const bottom = 1618;

function wrap(ctx, value, width) {
  return value.split('\n').flatMap(paragraph => {
    const lines = [];
    let line = '';
    for (const char of paragraph) {
      if (line && ctx.measureText(line + char).width > width) {
        lines.push(line);
        line = '';
      }
      line += char;
    }
    return [...lines, line];
  });
}

export async function createReport({ title, filename, cover, requests, pageCount, getImage, onProgress }) {
  const date = new Intl.DateTimeFormat('ko-KR', {dateStyle:'long'}).format(new Date());
  const glyphs = [title, filename, date, cover.project, cover.recipient, cover.sender, cover.note,
    '수정요청서 원본 페이지 요청 항목 계속 수신 발신 작성일 검토 자료 전달 사항 내용 미입력 등록된 없음',
    ...requests.map(request => request.text)].join(' ');
  await document.fonts.load('24px "Noto Sans KR"', glyphs);
  await document.fonts.ready;

  const pdf = new jsPDF({unit:'mm',format:'a4',compress:true});
  pdf.setProperties({title, subject:'수정요청서'});
  const canvas = document.createElement('canvas');
  canvas.width = 1240;
  canvas.height = 1754;
  const ctx = canvas.getContext('2d');
  let outputCount = 0;
  let y;

  function fill(color, x, y, w, h) {
    ctx.fillStyle = color;
    ctx.fillRect(x, y, w, h);
  }
  function text(value, x, y, size = 24, color = navy, width) {
    ctx.font = `${size}px "Noto Sans KR"`;
    ctx.fillStyle = color;
    if (width) {
      let clipped = value;
      while (clipped && ctx.measureText(clipped + (clipped !== value ? '…' : '')).width > width) {
        clipped = clipped.slice(0,-1);
      }
      value = clipped + (clipped !== value ? '…' : '');
    }
    ctx.fillText(value, x, y);
  }
  function paper() {
    fill('#ffffff',0,0,canvas.width,canvas.height);
    fill(navy,0,0,canvas.width,12);
  }
  function flush() {
    fill(border,left,1660,right-left,2);
    text(title,left,1704,18,muted,880);
    ctx.textAlign = 'right';
    text(String(++outputCount).padStart(2,'0'),right,1704,19,muted);
    ctx.textAlign = 'left';
    if (outputCount > 1) pdf.addPage();
    pdf.addImage(canvas.toDataURL('image/jpeg',.96),'JPEG',0,0,210,297);
  }
  function coverPage() {
    paper();
    fill('#f2f6fc',0,12,1240,610);
    if (cover.logo) {
      const scale = Math.min(300 / cover.logo.width, 116 / cover.logo.height);
      ctx.drawImage(cover.logo,left,88,cover.logo.width*scale,cover.logo.height*scale);
    }
    fill(blue,left,248,66,6);
    ctx.font = '63px "Noto Sans KR"';
    const headline = wrap(ctx,title,right-left);
    headline.forEach((line,index) => text(line,left,351+index*82,63));
    let row = 370 + headline.length*82;
    if (cover.project) {
      ctx.font = '30px "Noto Sans KR"';
      for (const line of wrap(ctx,cover.project,right-left)) {
        text(line,left,row,30,blue);
        row += 45;
      }
    }
    row = Math.max(675,row+44);
    for (const [label,value] of [['수신',cover.recipient],['발신',cover.sender],['작성일',date],['검토 자료',filename]]) {
      if (!value) continue;
      fill(border,left,row-34,right-left,1);
      text(label,left,row,21,muted);
      text(value,left+158,row,24,navy,right-left-158);
      row += 62;
    }
    if (cover.note) {
      // Compact repeated blank lines so the optional note fits on a single cover.
      const note = cover.note.replace(/\n[\t ]*\n(?:[\t ]*\n)*/g,'\n\n');
      ctx.font = '24px "Noto Sans KR"';
      const noteWidth = right-left-64;
      const availableLines = Math.floor((bottom-row-77)/35)+1;
      let noteLines = wrap(ctx,note,noteWidth);
      if (noteLines.length > availableLines) noteLines = wrap(ctx,note.replace(/\s+/g,' '),noteWidth);
      fill('#f2f6fc',left,row-10,right-left,82+noteLines.length*35);
      text('전달 사항',left+32,row+35,21,blue);
      noteLines.forEach((line,index) => text(line,left+32,row+77+index*35,24));
    }
    flush();
  }
  function bodyHeader(index, continued = false) {
    paper();
    text(title,left,79,36,navy,right-left);
    text(`원본 ${index+1} / ${pageCount}페이지${continued ? ' · 수정 요청 계속' : ''}`,left,123,22,blue);
    if (!continued) text(filename,480,123,19,muted,right-480);
    fill(border,left,151,right-left,2);
    y = 187;
  }
  function marker(number,x,y,radius = 20) {
    ctx.beginPath();
    ctx.arc(x,y,radius,0,2*Math.PI);
    ctx.fillStyle = '#ba392e';
    ctx.fill();
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = 3;
    ctx.stroke();
    ctx.textAlign = 'center';
    text(String(number),x,y+7,number >= 100 ? 15 : 20,'#ffffff');
    ctx.textAlign = 'left';
  }

  if (cover.enabled) coverPage();
  for (let index=0; index<pageCount; index++) {
    onProgress(index);
    bodyHeader(index);
    const source = await getImage(index);
    const scale = Math.min((right-left)/source.width,820/source.height);
    const width = source.width*scale;
    const height = source.height*scale;
    const x = (canvas.width-width)/2;
    fill('#fff',x-1,y-1,width+2,height+2);
    ctx.drawImage(source,x,y,width,height);
    ctx.lineWidth = 1;
    ctx.strokeStyle = border;
    ctx.strokeRect(x,y,width,height);
    const entries = requests.filter(request => request.page === index);
    for (const request of entries) marker(request.id,x+width*request.x,y+height*request.y);

    y += height + 44;
    fill('#eef3fa',left,y,right-left,53);
    text('수정 요청',left+20,y+35,24,navy);
    ctx.textAlign = 'right';
    text(entries.length + '개',right-20,y+34,20,blue);
    ctx.textAlign = 'left';
    y += 93;
    if (!entries.length) text('등록된 수정 요청이 없습니다.',left+20,y,22,muted);
    for (const request of entries) {
      ctx.font = '24px "Noto Sans KR"';
      const lines = wrap(ctx,request.text.trim() || '요청 내용 미입력',right-left-96);
      for (const [lineIndex,line] of lines.entries()) {
        const continued = y > bottom;
        if (continued) {
          flush();
          bodyHeader(index,true);
          y += 25;
        }
        if (lineIndex === 0 || continued) marker(request.id,left+26,y-8,18);
        text(line,left+65,y,24);
        y += 38;
      }
      fill(border,left,y-17,right-left,1);
      y += 24;
    }
    flush();
  }
  const name = title.replace(/[\\/:*?"<>|\u0000-\u001f]/g,'-').replace(/[. ]+$/,'');
  pdf.save((name || '수정요청서')+'.pdf');
  return outputCount;
}
