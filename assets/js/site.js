/* ======== OWNER SETTINGS: paste your Apps Script web app URL (ends in /exec) ======== */
const CONFIG={SCRIPT_URL:'https://script.google.com/macros/s/AKfycbw8oYavjxLlB7SRvJMxLNSrATclL7PdlFjkpzMx5pEQ22_HEjhAahhuXyizA-pil3w/exec',WHATSAPP:'2348104141138',
  PAY_BANK:'YOUR BANK NAME, account name OKV Technology Consults, account number 0000000000'};
/* Price, bank details, Paystack and notices are controlled from your OKV admin dashboard once the URL above is set. */
/* ===================================================================================== */
const $=s=>document.querySelector(s);
const configured=()=>/^https:\/\/script\.google\.com\//.test(CONFIG.SCRIPT_URL);
document.querySelectorAll('.js-wa').forEach(a=>a.href='https://wa.me/'+CONFIG.WHATSAPP);
document.querySelectorAll('.js-bank').forEach(a=>a.textContent=CONFIG.PAY_BANK);
document.querySelectorAll('.js-year').forEach(a=>a.textContent=new Date().getFullYear());
let REMOTE={};

function jsonp(params){return new Promise((res,rej)=>{
  if(!configured())return rej(new Error('This page is not connected yet. Please WhatsApp +234 810 414 1138.'));
  const cb='okv'+Date.now()+Math.floor(Math.random()*1e6),s=document.createElement('script');
  const done=()=>{clearTimeout(t);delete window[cb];s.remove()};
  const t=setTimeout(()=>{done();rej(new Error('The server took too long. Please try again.'))},30000);
  window[cb]=d=>{done();d&&d.ok?res(d):rej(new Error((d&&d.error)||'Submission failed'))};
  s.onerror=()=>{done();rej(new Error('Network error. Check your internet and try again.'))};
  s.src=CONFIG.SCRIPT_URL+'?'+new URLSearchParams(Object.assign({},params,{callback:cb})).toString();document.head.appendChild(s)})}
const formData=f=>Object.fromEntries(new FormData(f).entries());
const fileB64=f=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result).split(',')[1]);r.onerror=()=>rej(new Error('Could not read the file'));r.readAsDataURL(f)});

/* ---- values controlled from the OKV admin dashboard ---- */
function applyConfig(c){REMOTE=c;
  if(c.price)document.querySelectorAll('.js-price').forEach(e=>e.textContent=Number(c.price).toLocaleString('en-US'));
  if(c.price)document.querySelectorAll('input[name=amount]').forEach(e=>e.value=c.price);
  if(c.pay&&c.pay.bank){const t=c.pay.bank+', account name '+c.pay.account_name+', account number '+c.pay.account_no+(c.pay.instructions?'. '+c.pay.instructions:'');document.querySelectorAll('.js-bank').forEach(e=>e.textContent=t)}
  const manual=c.manual_enabled!==false,ps=!!c.paystack;
  document.querySelectorAll('.manual-only').forEach(e=>e.classList.toggle('d-none',!manual));
  document.querySelectorAll('.ps-only,#psBox').forEach(e=>e.classList.toggle('d-none',!ps));
  document.querySelectorAll('.ps-or').forEach(e=>e.classList.toggle('d-none',!(ps&&manual)));document.querySelectorAll('.ps-not').forEach(e=>e.classList.toggle('d-none',ps&&manual));
  const np=$('#nopay');if(np)np.classList.toggle('d-none',manual||ps);
  if(c.announcement){const bar=document.createElement('div');bar.className='text-center small py-2 px-3';bar.style.cssText='background:#fef3c7;color:#78350f';bar.textContent=c.announcement;document.body.insertBefore(bar,document.body.firstChild)}
  if(c.install&&c.install.text){document.querySelectorAll('.js-install-text').forEach(e=>e.textContent=c.install.text)}}
if(configured())jsonp({action:'config'}).then(applyConfig).catch(()=>{});

/* ---- cookie notice (only essential storage is used) ---- */
(function(){const bar=$('#cookiebar');if(!bar)return;let ok=false;try{ok=localStorage.getItem('okv_consent')==='1'}catch(e){}
  if(!ok)bar.classList.remove('d-none');
  $('#cookieok').onclick=()=>{try{localStorage.setItem('okv_consent','1')}catch(e){}bar.classList.add('d-none')}})();

/* ---- forms ---- */
const tf=$('#tf');if(tf)tf.onsubmit=async e=>{e.preventDefault();const b=$('#tbtn'),er=$('#terr');er.textContent='';b.disabled=true;b.textContent='Please wait...';
  try{const r=await jsonp(Object.assign(formData(e.target),{action:'trial'}));$('#tlink').href=r.downloadUrl;
    const show=(id,u)=>{const a=$(id);if(a&&u){a.href=u;a.classList.remove('d-none')}};show('#tlink',r.downloadUrl);show('#tlinkArm',r.macArmUrl);show('#tlinkIntel',r.macIntelUrl);
    const ua=navigator.userAgent||'';const mac=/Mac/i.test(ua)&&!/iPhone|iPad/i.test(ua);const order=mac?['#tlinkArm','#tlinkIntel','#tlink']:['#tlink','#tlinkArm','#tlinkIntel'];
    const box=$('#tlink').parentNode;order.forEach(id=>box.appendChild($(id)));
    if(r.pending){$('#tdone').innerHTML='<h2 class="h5">Thank you</h2><p>We have received your request. We will email your download link shortly.</p>'}
    $('#tform').classList.add('d-none');$('#tdone').classList.remove('d-none')}
  catch(x){er.textContent=x.message}finally{b.disabled=false;b.textContent='Get my free download'}};

const qf=$('#qf');if(qf)qf.onsubmit=async e=>{e.preventDefault();const m=$('#qmsg');m.className='mt-2 small text-muted';m.textContent='Sending...';
  try{await jsonp(Object.assign(formData(e.target),{action:'contact'}));m.className='mt-2 small text-success';m.textContent='Thank you! We received your message and will reply shortly.';e.target.reset()}
  catch(x){m.className='mt-2 small text-danger';m.textContent=x.message}};

const inf=$('#inf');if(inf)inf.onsubmit=async e=>{e.preventDefault();const b=$('#ibtn'),er=$('#ierr');er.textContent='';b.disabled=true;b.textContent='Sending...';
  try{const d=formData(e.target);const message='INSTALLATION REQUEST | Mode: '+d.mode+' | Computers: '+d.computers+' | Preferred date: '+(d.date||'any')+' | Notes: '+(d.notes||'');
    await jsonp({action:'install_request',name:d.name,hospital:d.hospital,email:d.email,phone:d.phone,city:d.city,message,hp:d.hp});
    $('#iform').classList.add('d-none');$('#idone').classList.remove('d-none')}
  catch(x){er.textContent=x.message}finally{b.disabled=false;b.textContent='Request installation'}};

const pf=$('#pf');if(pf)pf.onsubmit=async e=>{e.preventDefault();const b=$('#pbtn'),er=$('#perr');er.textContent='';b.disabled=true;b.textContent='Opening Paystack...';
  try{const r=await jsonp(Object.assign(formData(e.target),{action:'paystack_init'}));window.location.href=r.url}
  catch(x){er.textContent=x.message;b.disabled=false;b.textContent='Pay with Paystack'}};

const bf=$('#bf');if(bf)bf.onsubmit=async e=>{e.preventDefault();const b=$('#bbtn'),er=$('#berr');er.textContent='';
  const f=$('#bfile').files[0];if(!f){er.textContent='Attach your payment receipt';return}
  if(f.size>3*1024*1024){er.textContent='Receipt is too large (maximum 3 MB)';return}
  if(!configured()){er.textContent='This page is not connected yet. Please WhatsApp +234 810 414 1138.';return}
  b.disabled=true;b.textContent='Submitting...';
  try{const body=Object.assign(formData(e.target),{action:'purchase',receipt:await fileB64(f),receipt_name:f.name,receipt_mime:f.type});
    const ctl=new AbortController(),t=setTimeout(()=>ctl.abort(),60000);
    const r=await fetch(CONFIG.SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(body),signal:ctl.signal});clearTimeout(t);
    let d={};try{d=await r.json()}catch(_){d={ok:true}}   // Google may hide the reply from a browser; the submission is still recorded
    if(d&&d.ok===false)throw new Error(d.error||'Submission failed');
    $('#bform').classList.add('d-none');$('#bdone').classList.remove('d-none')}
  catch(x){er.textContent=x.name==='AbortError'?'The server took too long. Please try again.':x.message}finally{b.disabled=false;b.textContent='Submit payment and request code'}};

/* ---- arrive from the software's Buy licence button: ?install_id=HMS-XXXX&buy=1 ---- */
(function(){const q=new URLSearchParams(location.search),iid=(q.get('install_id')||'').toUpperCase();
  if(/^HMS-[0-9A-F]{8}$/.test(iid)){document.querySelectorAll('input[name=install_id]').forEach(e=>e.value=iid);
    const bb=$('#buyBanner');if(bb){bb.textContent='You are buying a licence for Installation ID '+iid+'. It will be attached to your order automatically.';bb.classList.remove('d-none')}}
  if(q.get('buy')==='1'){window.addEventListener('load',()=>{const m=$('#buyModal');if(m&&window.bootstrap)new bootstrap.Modal(m).show()})}})();
