import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
pdfjs.GlobalWorkerOptions.workerSrc=workerUrl;
export async function inspect(bytes){const d=await pdfjs.getDocument({data:new Uint8Array(bytes),isEvalSupported:false}).promise;const results=[];for(let i=1;i<=d.numPages;i++){const p=await d.getPage(i);const v=p.getViewport({scale:1.2});const c=document.createElement('canvas');c.width=v.width;c.height=v.height;await p.render({canvasContext:c.getContext('2d'),viewport:v}).promise;results.push(c.toDataURL('image/png').split(',')[1]);}await d.cleanup();return results;}
