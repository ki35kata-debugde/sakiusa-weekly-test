(function(){
const DECK="target1200",DECK_TEST="deck-"+DECK,DECK_DATE="2026-01-01",SUBJ="単語帳";
const css=document.createElement("link");css.rel="stylesheet";css.href="vocab.css";document.head.appendChild(css);
const pad=n=>String(n).padStart(4,"0");
const vq=()=>data.questions.filter(q=>q.deck===DECK).sort((a,b)=>Number(a.no)-Number(b.no));
const latestOf=id=>[...data.results].reverse().find(r=>r.qid===id);

// 単語帳の問題を毎週テストの一覧・予定に混ぜない
const baseTests=tests;tests=function(subject){return baseTests(subject).filter(t=>!String(t.id).startsWith("deck-"))};

// ---- ターゲット1200形式（8列）のExcel取込 ----
const HEAD_WORD="英単語・英熟語",HEAD_CLOZE="穴埋め問題";
async function parseTarget(file){
 if(!/\.xlsx?$|\.xlsm$/i.test(file.name))return[];
 const wb=XLSX.read(await file.arrayBuffer(),{type:"array"}),items=[];
 for(const name of wb.SheetNames){
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:""});
  const hi=rows.findIndex(r=>r.includes(HEAD_WORD)&&r.includes(HEAD_CLOZE));if(hi<0)continue;
  const h=rows[hi],c=label=>h.indexOf(label);
  const cNo=h.findIndex(x=>/^No\.?$/i.test(String(x).trim())),cWord=c(HEAD_WORD),cMean=c("日本語訳"),cSent=c("英語例文"),cSentJa=c("例文の日本語訳"),cCloze=c(HEAD_CLOZE),cCom=c("一言コメント"),cW=[c("誤答1"),c("誤答2"),c("誤答3")];
  for(const r of rows.slice(hi+1)){
   const no=Number(r[cNo]),word=String(r[cWord]||"").trim();if(!no||!word)continue;
   items.push({no,range:String(name),word,mean:String(r[cMean]||"").trim(),sent:String(r[cSent]||"").trim(),sentJa:String(r[cSentJa]||"").trim(),cloze:String(r[cCloze]||"").trim(),com:String(r[cCom]||"").trim(),wrong:cW.filter(i=>i>=0).map(i=>String(r[i]||"").trim()).filter(Boolean)});
  }
 }
 return items;
}
function upsertTarget(items){const now=new Date().toISOString();let added=0,updated=0;
 for(const it of items){
  const id=`v-${it.no}`,q=data.questions.find(x=>x.id===id);
  const fresh={id,deck:DECK,testId:DECK_TEST,date:DECK_DATE,subject:SUBJ,no:String(it.no),title:"ターゲット1200",range:japaneseGlyphs(it.range),prompt:japaneseGlyphs(it.mean),promptSub:it.cloze,promptJa:japaneseGlyphs(it.sentJa),answer:it.word,sentence:it.sent,comment:japaneseGlyphs(it.com),explanation:japaneseGlyphs(it.com),wrong:it.wrong||[],note:"",audioNo:pad(it.no),createdAt:q?.createdAt||now,updatedAt:now};
  if(q){Object.assign(q,fresh);updated++}else{data.questions.push(fresh);added++}
 }
 return{added,updated}}
const baseImport=importSheet;importSheet=async function(file){
 let items=[];try{items=await parseTarget(file)}catch(e){console.error(e)}
 if(!items.length)return baseImport(file);
 const {added,updated}=upsertTarget(items);
 data.importNotices=data.importNotices||[];data.importNotices.unshift({at:new Date().toISOString(),text:`単語帳更新：追加${added}語・上書き${updated}語`});data.importNotices=data.importNotices.slice(0,20);
 save();$("importMessage").textContent=`単語帳（ターゲット1200）として ${added}語を追加、${updated}語を更新しました。`;
};

// ---- 音声zip取込（共通ハンドラ） ----
async function refreshAudioStatus(){const el=$("vocabAudioStatus");if(!el)return;try{const n=await sakiAudio.count();el.textContent=n?`音声：${n}語ぶん取込済み`:"音声：未取込（音なしでも使えます）"}catch{el.textContent=""}}
function bindAudioImport(inputId,msgId){const input=$(inputId);if(!input)return;input.onchange=async e=>{const f=e.target.files[0];if(!f)return;const msg=$(msgId);msg.textContent="音声を取り込んでいます…";try{const n=await sakiAudio.importZip(f);msg.textContent=`音声${n}ファイルを取り込みました。`;refreshAudioStatus()}catch(x){msg.textContent=`読み込めません：${x.message}`}input.value=""}}

// ---- 保護者画面：単語帳の習得状況 ----
if(document.body.classList.contains("parent-site")){
 $("importFile").closest(".card").insertAdjacentHTML("beforeend",'<label class="secondary">音声zipを取り込む<input id="audioZipParent" type="file" accept=".zip" hidden></label><p id="audioMsgParent"></p>');
 bindAudioImport("audioZipParent","audioMsgParent");
 const baseHist=renderHistory;renderHistory=function(){baseHist();const qs=vq();if(!qs.length)return;
  const tried=qs.filter(q=>latestOf(q.id)),ok=tried.filter(q=>latestOf(q.id).good),weak=tried.filter(q=>!latestOf(q.id).good);
  $("history").insertAdjacentHTML("beforeend",`<div class="card"><h2>単語撃退（ターゲット1200）</h2><p>登録 ${qs.length}語・挑戦済み ${tried.length}語・いま覚えている ${ok.length}語</p>${weak.length?`<p><b>にがてな単語：</b>${weak.slice(0,12).map(q=>`${esc(q.answer)}（No.${esc(q.no)}）`).join("、")}${weak.length>12?` ほか${weak.length-12}語`:""}</p>`:tried.length?"<p>いま覚えていない単語はありません。すばらしい！</p>":""}</div>`)};
 render();
}

// ---- 子供画面：単語撃退 ----
if(!document.body.classList.contains("child-site"))return;
let vs=null;

function ranges(){const m=new Map;vq().forEach(q=>{if(!m.has(q.range))m.set(q.range,0);m.set(q.range,m.get(q.range)+1)});return[...m.entries()]}
function pool(){const range=$("vocabRange").value,from=Number($("vocabFrom").value)||0,to=Number($("vocabTo").value)||Infinity;
 return vq().filter(q=>(range==="all"||q.range===range)&&Number(q.no)>=from&&Number(q.no)<=to)}
function updateVocabCount(){const el=$("vocabCount");if(el)el.textContent=`単語：現在 ${pool().length}語選択中`}

function renderVocabSetup(){const box=$("vocabSetup");if(!box)return;const rs=ranges();
 if(!rs.length){box.innerHTML='<h2>単語撃退</h2><p>単語帳がまだ登録されていません。保護者用サイトの「問題を取り込む」から、ターゲット1200のExcelを取り込んでもらいましょう。</p>';return}
 box.innerHTML=`<h2>単語撃退の準備</h2><p id="vocabAudioStatus" class="v-status"></p><div class="v-setup-grid">
  <label>はんい<select id="vocabRange"><option value="all">すべて</option>${rs.map(([r,n])=>`<option value="${esc(r)}">${esc(r)}（${n}語）</option>`).join("")}</select></label>
  <label>No.でしぼる（省略可）<span class="v-nos"><input id="vocabFrom" type="number" inputmode="numeric" placeholder="はじめ"><span>〜</span><input id="vocabTo" type="number" inputmode="numeric" placeholder="おわり"></span></label>
  <label>やり方<select id="vocabMode"><option value="choice">4たくバトル</option><option value="card">カードめくり</option></select></label>
  <label>出す順番<select id="vocabOrder"><option value="smart">おぼえてない順</option><option value="no">No.順</option><option value="random">ランダム</option></select></label>
  <label>問題数<select id="vocabLimit"><option value="10">10語</option><option value="20">20語</option><option value="all">すべて</option></select></label>
  <div id="vocabCount" class="v-status"></div>
  <button id="vocabStart" class="primary">たたかう！</button>
  <label class="secondary" style="text-align:center">音声zipを取り込む<input id="audioZipChild" type="file" accept=".zip" hidden></label><p id="audioMsgChild" class="v-status"></p>
 </div>`;
 ["vocabRange","vocabFrom","vocabTo"].forEach(id=>$(id).addEventListener("input",updateVocabCount));
 $("vocabStart").onclick=vStart;bindAudioImport("audioZipChild","audioMsgChild");updateVocabCount();refreshAudioStatus()}

function vPriority(q){const rs=data.results.filter(r=>r.qid===q.id);return[rs.length===0?0:rs.some(r=>!r.good)?1:2,rs.length,Math.random()]}
function vStart(){let qs=pool();if(!qs.length)return alert("この条件で出せる単語がありません。");
 const order=$("vocabOrder").value;
 if(order==="smart")qs.sort((a,b)=>{const x=vPriority(a),y=vPriority(b);return x[0]-y[0]||x[1]-y[1]||x[2]-y[2]});
 else if(order==="random")qs=shuffle(qs);
 const limit=$("vocabLimit").value;if(limit!=="all")qs=qs.slice(0,Number(limit));
 vs={id:`a-${Date.now()}`,mode:$("vocabMode").value,queue:qs.map(q=>q.id),poolIds:pool().map(q=>q.id),index:0,records:[],started:0,pendingRetry:[]};
 $("vocabSetup").hidden=true;$("vocabFinish").hidden=true;$("vocabQuestion").hidden=false;vAsk()}

function frontHtml(q){return`<div class="v-front"><div class="v-meaning">${esc(q.prompt)}</div>${q.promptSub?`<div class="v-sub">${esc(q.promptSub)}</div>`:""}${q.promptJa?`<div class="v-sub-ja">${esc(q.promptJa)}</div>`:""}</div>`}
function backHtml(q){return`<div class="v-back"><div class="v-word">${esc(q.answer)}</div>${q.sentence?`<div class="v-sub">${esc(q.sentence)}</div>`:""}${q.promptJa?`<div class="v-sub-ja">${esc(q.promptJa)}</div>`:""}${q.comment?`<div class="v-comment">${esc(q.comment)}</div>`:""}</div>`}

function vAsk(){const q=data.questions.find(x=>x.id===vs.queue[vs.index]);vs.started=performance.now();$("mascot").src="assets/sakiusa-idle.gif";
 $("vocabQuestion").innerHTML=`<div class="quiz-head"><span class="v-progress">${vs.index+1} / ${vs.queue.length}</span><button id="vocabQuit" class="tiny">🥕 撤退！（途中結果へ）</button></div><p class="v-status">No.${esc(q.no)}・${esc(q.range)}</p><div id="vocabBody"></div>`;
 $("vocabQuit").onclick=()=>vFinish(true);
 if(vs.mode==="card"){
  $("vocabBody").innerHTML=`<div class="v-flip-card" id="vocabFlip">${frontHtml(q)}<div class="v-flip-hint">タップでめくる</div></div>`;
  $("vocabFlip").onclick=()=>{sakiAudio.play(q.audioNo);$("vocabBody").innerHTML=`<div class="v-flip-card">${backHtml(q)}</div><div class="v-self-buttons"><button data-self="correct" class="primary">○ おぼえてた</button><button data-self="wrong" class="secondary">× まだ</button><button data-self="later" class="secondary">△ あとで</button></div>`;
   document.querySelectorAll("#vocabBody [data-self]").forEach(b=>b.onclick=()=>{vRecord(q,b.dataset.self==="correct",b.dataset.self,b.dataset.self!=="correct");vNext()})};
 }else{
  const src=vs.poolIds&&vs.poolIds.length>=4?vq().filter(x=>vs.poolIds.includes(x.id)):vq(),
   picked=[...new Set((q.wrong||[]).filter(a=>a&&a!==q.answer))].slice(0,3),
   cands=[...new Set(src.filter(x=>x.id!==q.id).map(x=>x.answer))].filter(a=>a!==q.answer&&!picked.includes(a)),
   extra=[...new Set(vq().map(x=>x.answer))].filter(a=>a!==q.answer&&!picked.includes(a)&&!cands.includes(a)),
   fill=shuffle(cands);
  while(picked.length<3&&fill.length)picked.push(fill.pop());
  while(picked.length<3&&extra.length)picked.push(extra.splice(Math.floor(Math.random()*extra.length),1)[0]);
  const ops=shuffle([q.answer,...picked].slice(0,4));
  $("vocabBody").innerHTML=`${frontHtml(q)}<div class="choices">${ops.map(x=>`<button class="choice" data-answer="${esc(x)}">${esc(x)}</button>`).join("")}</div><div id="vocabFeedback"></div>`;
  document.querySelectorAll("#vocabBody .choice").forEach(b=>b.onclick=()=>{const elapsed=performance.now()-vs.started,good=b.dataset.answer===q.answer;
   document.querySelectorAll("#vocabBody .choice").forEach(x=>{x.disabled=true;if(x.dataset.answer===q.answer)x.classList.add("correct")});if(!good)b.classList.add("wrong");
   sakiAudio.play(q.audioNo);vRecord(q,good,b.dataset.answer,!good,elapsed);
   $("vocabFeedback").innerHTML=`<div class="v-feedback ${good?"good":"bad"}"><h3>${good?"正解！":"おしい！"}</h3>${backHtml(q)}</div><button id="vocabNext" class="primary">次へ</button>`;
   $("vocabNext").onclick=vNext;
   $("vocabNext").scrollIntoView({behavior:"smooth",block:"end"})});
 }}

function vRecord(q,good,user,retry,elapsed=0){
 data.results.push({id:`r-${Date.now()}-${Math.random()}`,attempt:vs.id,qid:q.id,testId:q.testId,subject:SUBJ,title:"単語撃退",prompt:q.prompt,answer:q.answer,explanation:q.comment||"",user,good,elapsed:Math.round(elapsed),at:new Date().toISOString()});
 vs.records.push({qid:q.id,good});
 if(retry&&!vs.pendingRetry.includes(q.id))vs.pendingRetry.push(q.id);
 localStorage.setItem(KEY,JSON.stringify(data));
 $("mascot").src=good?"assets/sakiusa-correct.gif":"assets/sakiusa-wrong.gif"}

function vNext(){vs.index++;
 if(vs.index>=vs.queue.length&&vs.pendingRetry.length){vs.queue.push(...vs.pendingRetry);vs.pendingRetry=[]}
 vs.index<vs.queue.length?vAsk():vFinish(false)}

function vFinish(early){const rs=vs.records,ok=rs.filter(r=>r.good).length,rate=rs.length?Math.round(ok/rs.length*100):0,
 wrong=[...new Set(rs.filter(r=>!r.good).map(r=>r.qid))].map(id=>data.questions.find(q=>q.id===id)).filter(Boolean),
 msg=rate===100?"全単語撃退！見事でおじゃ！":rate>=80?"あと少しで全撃退でおじゃ！":rate>=60?"いい調子。にがてをもう一度！":"ここから伸びるでおじゃ。少しずつ覚えよう！";
 $("vocabQuestion").hidden=true;$("vocabFinish").hidden=false;$("mascot").src="assets/sakiusa-complete.gif";
 $("vocabFinish").innerHTML=`<div class="center"><h2>${early?"途中結果":"単語撃退 完了！"}</h2><div class="big">${rate}%</div><p>今回 ${rs.length}問・正解 ${ok}問</p><h3>${msg}</h3></div><h3>まちがえた単語</h3>${wrong.length?wrong.map(q=>`<div class="wrong-item"><b>${esc(q.prompt)} → ${esc(q.answer)}</b><p>${esc(q.sentence||"")}</p><p>${esc(q.comment||"")}</p></div>`).join(""):"<p>ありません。すばらしい！</p>"}<button id="vocabHome" class="primary">じゅんびへもどる</button> ${wrong.length?'<button id="vocabRetryWrong" class="secondary">まちがいだけもう一度</button>':""}`;
 $("vocabHome").onclick=()=>{renderVocabSetup();$("vocabSetup").hidden=false;$("vocabFinish").hidden=true;render()};
 const rw=$("vocabRetryWrong");if(rw)rw.onclick=()=>{vs={id:`a-${Date.now()}`,mode:vs.mode,queue:wrong.map(q=>q.id),index:0,records:[],started:0,pendingRetry:[]};$("vocabFinish").hidden=true;$("vocabQuestion").hidden=false;vAsk()};
 localStorage.setItem(KEY,JSON.stringify(data));render()}

let setupSig="";function setupSignature(){return ranges().map(([r,n])=>`${r}:${n}`).join("|")}
const baseRender=render;render=function(){baseRender();const sig=setupSignature();if($("vocabSetup")&&!$("vocabSetup").hidden&&sig!==setupSig){setupSig=sig;renderVocabSetup()}};
setupSig=setupSignature();renderVocabSetup();
})();
