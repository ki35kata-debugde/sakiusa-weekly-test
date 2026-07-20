// ターゲット1200「単語のみ」音声(TG1200_1_*_XXXX_YYYY.mp3)をNo別に分割する
// 使い方: node tools/split-audio.mjs <入力フォルダ> <出力フォルダ>
//   ffmpeg は PATH か、環境変数 FFMPEG で指定
// 分割設定(検証済み): 無音 -38dB/1.5秒以上で区切り、各セグメントは
//   発話0.15秒前〜次の無音開始+0.3秒。先頭のセクションアナウンスは破棄。
import {execFile} from "node:child_process";
import {promisify} from "node:util";
import {readdirSync, mkdirSync, statSync} from "node:fs";
import {join} from "node:path";

const run = promisify(execFile);
const FFMPEG = process.env.FFMPEG || "ffmpeg";
const [,, inDir, outDir] = process.argv;
if (!inDir || !outDir) { console.error("usage: node split-audio.mjs <入力フォルダ> <出力フォルダ>"); process.exit(1); }
mkdirSync(outDir, {recursive: true});

function findMp3(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out.push(...findMp3(p));
    else if (/^TG1200_1_.*_\d{4}_\d{4}\.mp3$/i.test(name)) out.push(p);
  }
  return out;
}

async function speechBlocks(src) {
  const {stderr} = await run(FFMPEG, ["-i", src, "-af", "silencedetect=noise=-38dB:d=1.5", "-f", "null", "-"], {maxBuffer: 1e8});
  const starts = [...stderr.matchAll(/silence_start: ([\d.]+)/g)].map(m => Number(m[1]));
  const ends = [...stderr.matchAll(/silence_end: ([\d.]+)/g)].map(m => Number(m[1]));
  const dm = stderr.match(/Duration: (\d+):(\d+):([\d.]+)/);
  const dur = Number(dm[1]) * 3600 + Number(dm[2]) * 60 + Number(dm[3]);
  // 発話区間 = 無音の切れ目の間。0.35秒未満は末尾のノイズ断片とみなして捨てる
  const blocks = [];
  let pos = 0;
  for (let i = 0; i < starts.length; i++) { blocks.push([pos, starts[i]]); pos = ends[i] ?? dur; }
  blocks.push([pos, dur]);
  return blocks.filter(([s, e]) => e - s >= 0.35);
}

const files = findMp3(inDir).sort();
if (!files.length) { console.error("対象mp3が見つかりません: " + inDir); process.exit(1); }
console.log(`${files.length}ファイルを処理します`);
const errors = [];
let written = 0;

for (const src of files) {
  const m = src.match(/_(\d{4})_(\d{4})\.mp3$/i);
  const from = Number(m[1]), to = Number(m[2]), expected = to - from + 1;
  const blocks = await speechBlocks(src);
  let items;
  if (blocks.length === expected) items = blocks;               // アナウンスなし
  else if (blocks.length === expected + 1) items = blocks.slice(1); // 先頭=アナウンス
  else { errors.push(`${src} : 発話${blocks.length}個 / 期待${expected}個`); continue; }
  for (let i = 0; i < items.length; i++) {
    const no = String(from + i).padStart(4, "0");
    const [s, e] = items[i];
    const args = ["-y", "-loglevel", "error", "-i", src,
      "-ss", Math.max(0, s - 0.15).toFixed(3), "-to", (e + 0.3).toFixed(3),
      "-vn", "-map_metadata", "-1", "-c:a", "libmp3lame", "-q:a", "4", join(outDir, no + ".mp3")];
    await run(FFMPEG, args);
    written++;
  }
  console.log(`OK  No.${from}〜${to} (${expected}語)  ${src.split(/[\\/]/).pop()}`);
}

console.log(`\n完了: ${written}ファイル出力`);
if (errors.length) { console.log("★要確認(未出力):"); errors.forEach(e => console.log("  " + e)); process.exit(2); }
