(function(){
const DB="sakiusa-audio",STORE="clips";let dbp=null;
function db(){if(!dbp)dbp=new Promise((res,rej)=>{const r=indexedDB.open(DB,1);r.onupgradeneeded=()=>r.result.createObjectStore(STORE);r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)});return dbp}
async function put(entries){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(STORE,"readwrite"),st=tx.objectStore(STORE);entries.forEach(([k,v])=>st.put(v,k));tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
async function get(key){const d=await db();return new Promise((res,rej)=>{const rq=d.transaction(STORE).objectStore(STORE).get(key);rq.onsuccess=()=>res(rq.result||null);rq.onerror=()=>rej(rq.error)})}
async function count(){const d=await db();return new Promise((res,rej)=>{const rq=d.transaction(STORE).objectStore(STORE).count();rq.onsuccess=()=>res(rq.result);rq.onerror=()=>rej(rq.error)})}
const urls=new Map;let player=null;
async function play(no){try{if(!no)return false;const key=String(no).padStart(4,"0");let u=urls.get(key);if(!u){const b=await get(key);if(!b)return false;u=URL.createObjectURL(b);urls.set(key,u)}if(player)player.pause();player=new Audio(u);await player.play();return true}catch{return false}}
async function importZip(file){const zip=await JSZip.loadAsync(await file.arrayBuffer()),entries=[];for(const name of Object.keys(zip.files)){const f=zip.files[name];if(f.dir)continue;const m=name.match(/(\d{4})\.mp3$/);if(!m)continue;entries.push([m[1],new Blob([await f.async("arraybuffer")],{type:"audio/mpeg"})])}if(!entries.length)throw Error("zipの中に「0001.mp3」形式のファイルが見つかりません。");await put(entries);return entries.length}
window.sakiAudio={play,count,importZip};
})();
