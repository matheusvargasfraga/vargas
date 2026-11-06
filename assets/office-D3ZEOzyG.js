const i=(t,e)=>t[e]|t[e+1]<<8,x=(t,e)=>(t[e]|t[e+1]<<8|t[e+2]<<16|t[e+3]<<24)>>>0,S=new TextDecoder;function j(t){let e=-1;for(let r=t.length-22;r>=Math.max(0,t.length-66e3);r--)if(x(t,r)===101010256){e=r;break}if(e<0)throw new Error("não é um arquivo do Office válido");const o=i(t,e+10);let n=x(t,e+16);const s=new Map;for(let r=0;r<o&&x(t,n)===33639248;r++){const c=i(t,n+10),a=x(t,n+20),m=i(t,n+28),k=i(t,n+30),p=i(t,n+32),w=x(t,n+42),f=S.decode(t.subarray(n+46,n+46+m));s.set(f,async()=>{const l=w+30+i(t,w+26)+i(t,w+28),h=t.subarray(l,l+a);if(c===0)return S.decode(h);if(c!==8)throw new Error("compressão não suportada");const u=new Blob([h]).stream().pipeThrough(new DecompressionStream("deflate-raw"));return S.decode(new Uint8Array(await new Response(u).arrayBuffer()))}),n+=46+m+k+p}return s}const q={amp:"&",lt:"<",gt:">",quot:'"',apos:"'",nbsp:" "},A=t=>t.replace(/&(#x?[0-9a-f]+|\w+);/gi,(e,o)=>o[0]==="#"?String.fromCodePoint(o[1]==="x"||o[1]==="X"?parseInt(o.slice(2),16):+o.slice(1)):q[o]??e),d=t=>A(t.replace(/<[^>]+>/g,"")),g=t=>+(t.match(/(\d+)\.xml$/)||[0,0])[1],P=t=>{let e=0;for(const o of t.replace(/\d+/g,""))e=e*26+o.charCodeAt(0)-64;return e-1};async function T(t){const e=await t.get("word/document.xml")();return d(e.replace(/<w:tab\/>/g,"	").replace(/<w:br[^>]*\/>/g,`
`).replace(/<\/w:p>/g,`
`).replace(/<\/w:tc>/g," | ").replace(/<\/w:tr>/g,`
`))}async function C(t){const e=[...t.keys()].filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((n,s)=>g(n)-g(s)),o=[];for(const n of e){const s=await t.get(n)();o.push("--- Slide "+g(n)+` ---
`+d(s.replace(/<\/a:p>/g,`
`)).replace(/\n{2,}/g,`
`).trim());const r=t.get("ppt/notesSlides/notesSlide"+g(n)+".xml");if(r){const c=d((await r()).replace(/<\/a:p>/g,`
`)).replace(/\n{2,}/g,`
`).trim();c&&!/^\d+$/.test(c)&&o.push("(notas: "+c+")")}}return o.join(`

`)}async function D(t,e=2e3){const o=t.get("xl/sharedStrings.xml")?[...(await t.get("xl/sharedStrings.xml")()).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(a=>d(a[1].replace(/<rPh[\s\S]*?<\/rPh>/g,""))):[],s=[...(t.get("xl/workbook.xml")?await t.get("xl/workbook.xml")():"").matchAll(/<sheet [^>]*name="([^"]*)"/g)].map(a=>A(a[1])),r=[...t.keys()].filter(a=>/^xl\/worksheets\/sheet\d+\.xml$/.test(a)).sort((a,m)=>g(a)-g(m)),c=[];for(const[a,m]of r.entries()){const k=await t.get(m)(),p=[];for(const w of k.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)){const f=[];for(const l of w[1].matchAll(/<c ([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){const h=l[1],u=l[2]||"",v=(h.match(/r="([A-Z]+\d+)"/)||[])[1],E=(h.match(/t="(\w+)"/)||[])[1],y=(u.match(/<v>([\s\S]*?)<\/v>/)||[])[1],$=E==="s"?o[+y]??"":E==="inlineStr"?d(u):y!=null?A(y):"";f[v?P(v):f.length]=String($).replace(/[\n;]/g," ")}if(f.some(l=>l)&&p.push(Array.from(f,l=>l??"").join(";").replace(/;+$/,"")),p.length>=e){p.push("… (planilha cortada em "+e+" linhas)");break}}c.push("--- Aba: "+(s[a]||"Planilha "+(a+1))+` ---
`+p.join(`
`))}return c.join(`

`)}async function O(t){const e=await t.get("content.xml")();return d(e.replace(/<text:tab\/>/g,"	").replace(/<text:line-break\/>/g,`
`).replace(/<\/text:(p|h)>/g,`
`).replace(/<\/table:table-cell>/g,";").replace(/<\/table:table-row>/g,`
`))}async function U(t,e=""){const o=t instanceof Uint8Array?t:new Uint8Array(t),n=j(o);let s;if(n.get("word/document.xml"))s=await T(n);else if(n.get("xl/workbook.xml")||[...n.keys()].some(r=>r.startsWith("xl/worksheets/")))s=await D(n);else if([...n.keys()].some(r=>r.startsWith("ppt/slides/")))s=await C(n);else if(n.get("content.xml"))s=await O(n);else throw new Error((e||"arquivo")+": formato do Office não reconhecido");return s.replace(/[ \t]+\n/g,`
`).replace(/\n{3,}/g,`

`).trim()}export{U as textoOffice};
