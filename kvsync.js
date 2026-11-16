/* MDWINGS – đồng bộ KiotViet trong trình duyệt.
   Dùng (trong tab app, đăng nhập quản trị):
     await kvSync({d0:'2026-11-10', d1:'2026-11-16', urls:{DanhSachSanPham:'https://cdn2-...xlsx', ...}, dry:false})
   Kết quả cũng nằm ở window.__sync */
(function(G){
const DATECOLS=new Set(['Thời gian','Ngày tạo','Thời gian tạo','Thời gian (theo giao dịch)']);
const td=new TextDecoder();
async function inflate(u8){if(G.__inflate)return G.__inflate(u8);const ds=new DecompressionStream('deflate-raw');const s=new Blob([u8]).stream().pipeThrough(ds);return new Uint8Array(await new Response(s).arrayBuffer())}
async function unzip(buf){const u=new Uint8Array(buf),v=new DataView(buf),out={};let e=u.length-22;while(e>=0&&v.getUint32(e,true)!==0x06054b50)e--;if(e<0)throw Error('zip?');
 let p=v.getUint32(e+16,true);const n=v.getUint16(e+10,true);
 for(let i=0;i<n;i++){const m=v.getUint16(p+10,true),cs=v.getUint32(p+20,true),fl=v.getUint16(p+28,true),xl=v.getUint16(p+30,true),cl=v.getUint16(p+32,true),off=v.getUint32(p+42,true);
  const name=td.decode(u.subarray(p+46,p+46+fl));p+=46+fl+xl+cl;
  const ds=off+30+v.getUint16(off+26,true)+v.getUint16(off+28,true);const raw=u.subarray(ds,ds+cs);out[name]={m,raw}}
 return {get:async nm=>{const k=Object.keys(out).find(k=>k.replace(/^\//,'')===nm);if(!k)return null;const f=out[k];return td.decode(f.m===0?f.raw:await inflate(f.raw))},names:Object.keys(out)}}
const unx=s=>s.replace(/\r\n?/g,'\n').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(_,d)=>String.fromCharCode(+d)).replace(/&#x([0-9a-f]+);/gi,(_,d)=>String.fromCharCode(parseInt(d,16))).replace(/&amp;/g,'&');
const colIdx=r=>{const L=r.match(/^[A-Z]+/)[0];let n=0;for(const ch of L)n=n*26+ch.charCodeAt(0)-64;return n-1};
async function readXlsx(buf){const z=await unzip(buf);
 const ss=[];const sx=await z.get('xl/sharedStrings.xml');if(sx)for(const m of sx.matchAll(/<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>/g))ss.push(unx([...m[1].matchAll(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map(x=>x[1]).join('')));
 const sheet=z.names.map(n=>n.replace(/^\//,'')).filter(n=>/^xl\/worksheets\/[^/]+\.xml$/.test(n)).sort()[0];const x=await z.get(sheet);
 const rows=[];for(const rm of x.matchAll(/<(?:\w+:)?row\b[^>]*>([\s\S]*?)<\/(?:\w+:)?row>/g)){const row=[];
  for(const cm of rm[1].matchAll(/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)){const at=cm[1],body=cm[2]||'';const r=(at.match(/\br="([A-Z]+)\d+"/)||[])[1];const t=(at.match(/\bt="(\w+)"/)||[])[1]||'n';
   let val=null;const vm=body.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
   if(t==='inlineStr'){val=unx([...body.matchAll(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map(x=>x[1]).join(''))}
   else if(vm){const s=vm[1];if(t==='s')val=ss[+s];else if(t==='str'||t==='e')val=unx(s);else if(t==='b')val=s==='1';else val=s===''?null:+s}
   if(val==='')val=null;row[r?colIdx(r):row.length]=val}
  rows.push(row)}
 const h=rows[0]||[];return rows.slice(1).map(r=>{const o={};h.forEach((k,i)=>{let v=r[i]===undefined?null:r[i];if(DATECOLS.has(k)&&typeof v==='number'){const naive=Math.round((v-25569)*86400000);v={ms:naive-7*3600000,d:new Date(naive).toISOString().slice(0,10)}}o[k]=v});return o})}

const n=v=>+(v||0);const rid=c=>'kv_'+String(c).replace(/[^A-Za-z0-9_.-]/g,'_');const pid=s=>rid('SP_'+s);
function build(T,D0,D1){const {P,K,NCC,I,R,N,Q,L}=T;const inwin=t=>t&&t.d>=D0&&t.d<=D1;
 const col=(row,...ns)=>{for(const x of ns)if(x in row)return row[x];return null};
 const grp=(rows,key)=>{const g=new Map();for(const r of rows)if(r[key]){if(!g.has(r[key]))g.set(r[key],[]);g.get(r[key]).push(r)}return g};
 const pcost={};P.forEach(p=>pcost[p['Mã hàng']]=n(p['Giá vốn']));const kby={};K.forEach(k=>kby[k['Mã khách hàng']]=k);
 const lcost=new Map(),txcost={};for(const r of L){if(!r['Mã giao dịch'])continue;const k=r['Mã giao dịch']+'\u0001'+r['Mã hàng'];if(!lcost.has(k))lcost.set(k,[]);lcost.get(k).push(n(r['Giá vốn/SP']));txcost[r['Mã giao dịch']]=Math.abs(n(r['Tổng giá vốn (theo giao dịch)']))}
 const costs=(code,items)=>{const used={};for(const i of items){const k=code+'\u0001'+i.sku,l=lcost.get(k);used[k]=used[k]||0;i.cost=l?l[Math.min(used[k],l.length-1)]:(pcost[i.sku]||0);used[k]++}
  if(code in txcost&&items.length){const d=txcost[code]-items.reduce((a,i)=>a+i.qty*i.cost,0);if(Math.abs(d)>1e-6){let j=0,best=-Infinity;items.forEach((x,ix)=>{const s=x.qty?x.qty*x.cost:-1;if(s>best){best=s;j=ix}});if(items[j].qty)items[j].cost+=d/items[j].qty}}return items};
 const orders=[];for(const [code,rs] of grp(I,'Mã hóa đơn')){const r=rs[0];if(!inwin(r['Thời gian'])||r['Trạng thái']==='Đã hủy')continue;const cc=r['Mã khách hàng']||'';
  const items=rs.filter(x=>x['Mã hàng']).map(x=>({pid:pid(x['Mã hàng']),sku:x['Mã hàng'],name:x['Tên hàng'],qty:n(x['Số lượng']),price:n(x['Giá bán']),list:n(x['Đơn giá'])||n(x['Giá bán'])}));
  orders.push({id:rid(code),code,date:r['Thời gian'].d,createdAt:r['Thời gian'].ms,cid:cc?rid('KH_'+cc):'',cname:r['Tên khách hàng']||'Khách lẻ',ccode:cc,channel:r['Kênh bán']||'Bán trực tiếp',status:'Hoàn thành',items:costs(code,items),goods:n(r['Tổng tiền hàng']),discount:n(r['Giảm giá hóa đơn']),ship:0,total:n(r['Khách cần trả']),tracking:'',note:r['Ghi chú']||'',createdByName:r['Người bán']||''})}
 const ocode={};orders.forEach(o=>ocode[o.code]=o);
 const rets=[];for(const [code,rs] of grp(R,'Mã trả hàng')){const r=rs[0];if(!inwin(r['Thời gian'])||r['Trạng thái']==='Đã hủy')continue;const cc=r['Mã khách hàng']||'',oc=r['Mã hóa đơn']||'';
  const items=rs.filter(x=>x['Mã hàng']).map(x=>({pid:pid(x['Mã hàng']),sku:x['Mã hàng'],name:x['Tên hàng'],qty:n(x['Số lượng']),price:n(x['Giá nhập lại'])}));
  rets.push({id:rid(code),code,date:r['Thời gian'].d,createdAt:r['Thời gian'].ms,orderId:oc?rid(oc):'',orderCode:oc,cid:cc?rid('KH_'+cc):'',cname:r['Tên khách hàng']||'Khách lẻ',items:costs(code,items),total:n(r['Cần trả khách']),note:r['Ghi chú']||'',canceled:false,createdByName:r['Người tạo']||''})}
 const rcode={};rets.forEach(x=>rcode[x.code]=x);
 const purch=[];for(const [code,rs] of grp(N,'Mã nhập hàng')){const r=rs[0];if(!inwin(r['Thời gian'])||['Đã hủy','Phiếu tạm'].includes(r['Trạng thái']))continue;
  purch.push({id:rid(code),code,date:r['Thời gian'].d,createdAt:r['Thời gian'].ms,sid:rid(r['Mã nhà cung cấp']),sname:r['Tên nhà cung cấp'],scode:r['Mã nhà cung cấp'],items:rs.filter(x=>x['Mã hàng']).map(x=>({pid:pid(x['Mã hàng']),sku:x['Mã hàng'],name:x['Tên hàng'],qty:n(x['Số lượng']),price:n(x['Giá nhập'])})),discount:n(r['Giảm giá phiếu nhập']),total:n(r['Cần trả NCC']),note:r['Ghi chú']||'',canceled:false,createdByName:r['Người nhập']||''})}
 const pcode={};purch.forEach(p=>pcode[p.code]=p);
 const byname={};K.forEach(k=>{const nm=(k['Tên khách hàng']||'').trim();(byname[nm]=byname[nm]||[]).push(k['Mã khách hàng'])});
 const sups={};for(const s of NCC){const code=col(s,'Mã nhà cung cấp','Mã NCC');if(code)sups[code]={name:col(s,'Tên nhà cung cấp','Tên NCC')||code,phone:String(col(s,'Điện thoại')||''),debt:n(col(s,'Nợ cần trả hiện tại','Nợ hiện tại','Nợ cần trả'))}}
 const supname={};Object.entries(sups).forEach(([k,v])=>supname[v.name]=k);
 const cash=[];let fund_end=null;const extra=new Set();const strip=b=>b.replace(/\.\d{2}$/,'');
 for(const q of Q){if(!q['Mã phiếu']||!q['Thời gian']){if(q['Người nộp/nhận']==='Tồn quỹ')fund_end=n(q['Giá trị']);continue}
  if(!inwin(q['Thời gian']))continue;const code=q['Mã phiếu'],v=n(q['Giá trị']);
  const cat=(q['Loại thu chi']||'').replace('Phiếu thu ','Thu ').replace('Phiếu chi ','Chi ').trim()||(v>0?'Thu khác':'Chi khác');
  const c={id:rid('Q_'+code),refCode:code,type:v>0?'thu':'chi',account:'Tiền mặt',amount:Math.abs(v),category:cat,partnerType:'',partnerId:'',partnerName:q['Người nộp/nhận']||'',date:q['Thời gian'].d,createdAt:q['Thời gian'].ms,note:'',canceled:false};
  const base=code.slice(2);
  if(code.startsWith('TTHD')){const o=ocode[base]||ocode[strip(base)];if(o)Object.assign(c,{partnerType:'customer',partnerId:o.cid,partnerName:o.cname,orderId:o.id});
   else{const m=base.match(/^(HD[\d.]+?)(\.\d{2})?$/);const ks=byname[c.partnerName.trim()]||[];c.orderId=m?rid(m[1]):'';
    if(ks.length===1){Object.assign(c,{partnerType:'customer',partnerId:rid('KH_'+ks[0])});extra.add(ks[0])}else c.kqkd=false}}
  else if(code.startsWith('PCPN')){const p=pcode[base]||pcode[strip(base)];if(p)Object.assign(c,{partnerType:'supplier',partnerId:p.sid,partnerName:p.sname,purchaseId:p.id});else if(c.partnerName in supname)Object.assign(c,{partnerType:'supplier',partnerId:rid(supname[c.partnerName])})}
  else if(code.startsWith('TTTH')){const x=rcode[base]||rcode[strip(base)];if(x)Object.assign(c,{partnerType:'customer',partnerId:x.cid,partnerName:x.cname,returnId:x.id})}
  cash.push(c)}
 const cd={},sd={},st={};const add=(o,k,v)=>o[k]=(o[k]||0)+v;
 orders.forEach(o=>{add(cd,o.cid,o.total);o.items.forEach(i=>add(st,i.pid,-i.qty))});rets.forEach(x=>{add(cd,x.cid,-x.total);x.items.forEach(i=>add(st,i.pid,i.qty))});
 purch.forEach(p=>{add(sd,p.sid,p.total);p.items.forEach(i=>add(st,i.pid,i.qty))});
 cash.forEach(c=>{if(c.partnerType==='customer')add(cd,c.partnerId,c.type==='thu'?-c.amount:c.amount);if(c.partnerType==='supplier')add(sd,c.partnerId,c.type==='chi'?-c.amount:c.amount)});
 const used=new Set([...orders.filter(o=>o.ccode).map(o=>o.ccode),...rets.filter(x=>x.cid).map(x=>x.cid.slice(6)),...extra,...K.filter(k=>n(k['Nợ cần thu hiện tại'])!==0).map(k=>k['Mã khách hàng'])]);
 const customers=[];for(const cc of [...used].sort()){const k=kby[cc];if(!k)continue;const i=rid('KH_'+cc);
  customers.push({id:i,code:cc,name:k['Tên khách hàng']||cc,phone:String(k['Điện thoại']||''),address:k['Địa chỉ']||'',taxCode:k['MST (xuất HĐ)']||'',note:k['Ghi chú']||'',openDebtIfNew:Math.round(n(k['Nợ cần thu hiện tại'])-(cd[i]||0)),createdAt:k['Ngày tạo']?k['Ngày tạo'].ms:0,curDebt:n(k['Nợ cần thu hiện tại'])})}
 const suppliers=Object.entries(sups).map(([code,s])=>({id:rid(code),code,name:s.name,phone:s.phone,address:'',taxCode:'',note:'',openDebtIfNew:Math.round(s.debt-(sd[rid(code)]||0)),createdAt:0,curDebt:s.debt}));
 const products=P.map(p=>({id:pid(p['Mã hàng']),sku:p['Mã hàng'],name:p['Tên hàng'],group:p['Nhóm hàng(3 Cấp)']||'',brand:p['Thương hiệu']||'',unit:p['ĐVT']||'cái',price:n(p['Giá bán']),hidden:!p['Đang kinh doanh'],openCostIfNew:n(p['Giá vốn']),openStockIfNew:n(p['Tồn kho'])-(st[pid(p['Mã hàng'])]||0),curStock:n(p['Tồn kho']),createdAt:p['Thời gian tạo']?p['Thời gian tạo'].ms:0}));
 const daily={};orders.forEach(o=>{daily[o.date]=daily[o.date]||[0,0];daily[o.date][0]+=o.goods-o.discount});rets.forEach(x=>{daily[x.date]=daily[x.date]||[0,0];daily[x.date][1]+=x.total});
 Object.values(daily).forEach(a=>{a[0]=Math.round(a[0]);a[1]=Math.round(a[1])});
 return {window:[D0,D1],products,customers,suppliers,orders,returns:rets,purchases:purch,cash,expect:{fund:fund_end,custDebtTotal:K.reduce((a,k)=>a+n(k['Nợ cần thu hiện tại']),0),daily}}}

async function apply(J){const fs=firebase.firestore(),uid=firebase.auth().currentUser.uid,[d0,d1]=J.window,ops=[];const base={createdBy:uid,importedFrom:'KiotViet'};
 for(const p of J.products){const ex=S.products.get(p.id);const f={sku:p.sku,name:p.name,group:p.group,brand:p.brand,unit:p.unit,price:p.price,hidden:p.hidden};
  ops.push(ex?['products',p.id,f,true]:['products',p.id,{...f,openStock:p.openStockIfNew,openCost:p.openCostIfNew,minStock:0,createdAt:p.createdAt,...base,createdByName:'KiotViet'},false])}
 for(const [col,arr] of [['customers',J.customers],['suppliers',J.suppliers]])for(const c of arr){const ex=S[col].get(c.id);const f={code:c.code,name:c.name,phone:c.phone};
  ops.push(ex?[col,c.id,f,true]:[col,c.id,{...f,address:c.address,taxCode:c.taxCode,note:c.note,openDebt:c.openDebtIfNew,createdAt:c.createdAt,...base,createdByName:'KiotViet'},false])}
 const lastKq={};[...S.cash.values()].sort((a,b)=>(a.createdAt||0)-(b.createdAt||0)).forEach(c=>{if(!c.partnerType)lastKq[c.category]=c.kqkd!==false});
 const inc={};for(const col of ['orders','returns','purchases','cash']){inc[col]=new Set();for(const x of J[col]){const {id,...d}=x;inc[col].add(id);const ex=S[col].get(id);
  if(col==='cash'&&!d.partnerType&&d.kqkd===undefined)d.kqkd=ex&&ex.kqkd!==undefined?ex.kqkd:(lastKq[d.category]??true);
  ops.push([col,id,{...d,...base,createdByName:d.createdByName||'KiotViet'},true])}}
 let cancel=0;const canc=[];for(const col of ['orders','returns','purchases','cash'])for(const x of S[col].values()){if(!x.id.startsWith('kv_')||x.id==='kv_Q_DAUKY'||!(x.date>=d0&&x.date<=d1)||inc[col].has(x.id))continue;
  if(col==='orders'){if(x.status!=='Đã hủy'){ops.push([col,x.id,{status:'Đã hủy',canceledAt:Date.now()},true]);cancel++;canc.push(x.id)}}else if(!x.canceled){ops.push([col,x.id,{canceled:true},true]);cancel++;canc.push(x.id)}}
 const newIds={};for(const col of ['orders','returns','purchases','cash'])newIds[col]=J[col].filter(x=>!S[col].get(x.id)).map(x=>x.code||x.refCode);
 G.__sync={status:'writing',done:0,total:ops.length};
 for(let i=0;i<ops.length;i+=400){const b=fs.batch();ops.slice(i,i+400).forEach(([col,id,d,merge])=>b.set(fs.collection(col).doc(id),d,{merge}));await b.commit();G.__sync.done+=Math.min(400,ops.length-i)}
 await new Promise(r=>setTimeout(r,3000));derive();
 const bad=[];for(const p of J.products){const s=D.st[p.id]||0;if(Math.abs(s-p.curStock)>1e-6)bad.push('Tồn '+p.sku+': app '+s+' / KV '+p.curStock)}
 for(const c of J.customers){const v=D.cd[c.id]||0;if(Math.abs(v-c.curDebt)>1)bad.push('Nợ KH '+c.code+': app '+v+' / KV '+c.curDebt)}
 for(const c of J.suppliers){const v=D.sd[c.id]||0;if(Math.abs(v-c.curDebt)>1)bad.push('Nợ NCC '+c.code+': app '+v+' / KV '+c.curDebt)}
 const fund=Object.values(D.bal).reduce((a,b)=>a+b,0);if(J.expect.fund!=null&&Math.abs(fund-J.expect.fund)>1)bad.push('Quỹ: app '+fund+' / KV '+J.expect.fund);
 const cdt=Object.values(D.cd).reduce((a,b)=>a+b,0);if(Math.abs(cdt-J.expect.custDebtTotal)>1)bad.push('Tổng nợ KH: app '+cdt+' / KV '+J.expect.custDebtTotal);
 for(const [d,[rv,rt]] of Object.entries(J.expect.daily)){const os=vals('orders').filter(o=>o.status!=='Đã hủy'&&o.date===d),rs=vals('returns').filter(r=>!r.canceled&&r.date===d);const a=sum(os,o=>o.goods-o.discount),b=sum(rs,r=>r.total);if(Math.abs(a-rv)>1||Math.abs(b-rt)>1)bad.push('Doanh thu '+d+': app '+a+'/'+b+' / KV '+rv+'/'+rt)}
 return {status:'done',window:J.window,writes:ops.length,canceled:canc,newDocs:newIds,fund,custDebt:cdt,mismatches:bad.slice(0,40),mismatchCount:bad.length}}

const NEED={P:'DanhSachSanPham',K:'DanhSachKhachHang',NCC:'DanhSachNhaCungCap',I:'DanhSachChiTietHoaDon',R:'DanhSachChiTietTraHang',N:'DanhSachChiTietNhapHang',Q:'SoQuy',L:'BaoCaoBanHangTheoLoiNhuan'};
G.kvReadXlsx=readXlsx;G.kvBuild=build;
G.kvSync=async({d0,d1,urls,bufs,dry})=>{try{G.__sync={status:'fetching'};const T={};
  for(const [k,pre] of Object.entries(NEED)){let buf=bufs&&bufs[pre];if(!buf){const u=urls&&(urls[pre]||Object.entries(urls).find(([n])=>n.startsWith(pre))?.[1]);if(!u)throw Error('THIẾU FILE: '+pre);const r=await fetch(u);if(!r.ok)throw Error('Tải '+pre+' lỗi '+r.status);buf=await r.arrayBuffer()}T[k]=await readXlsx(buf)}
  G.__sync={status:'building'};const J=build(T,d0,d1);G.__kvJ=J;
  const cnt={};for(const k of ['products','customers','suppliers','orders','returns','purchases','cash'])cnt[k]=J[k].length;
  if(dry)return G.__sync={status:'dry',counts:cnt,expect:J.expect};
  const res=await apply(J);res.counts=cnt;return G.__sync=res}catch(e){return G.__sync={status:'error',err:String(e&&e.stack||e)}}};
})(typeof window!=='undefined'?window:globalThis);
