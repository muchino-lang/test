import { chromium } from 'playwright';
import fs from 'fs';
const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium'});
const ctx=await b.newContext({viewport:{width:1280,height:1000},permissions:['clipboard-read','clipboard-write']});
const p=await ctx.newPage(); const errs=[];
p.on('pageerror',e=>errs.push('PAGEERROR: '+e.message));
p.on('console',m=>{if(m.type()==='error')errs.push('CONSOLE: '+m.text())});
p.on('dialog',d=>d.accept());
const ok=(n,c)=>console.log((c?'✓':'✗')+' '+n);
await p.goto('file:///home/user/test/daily-report-system/index.html'); await p.waitForTimeout(400);
ok('スキル78項目', await p.evaluate(()=>SKILLS.length)===78);
ok('重点3項目', await p.locator('#focusGrid .fcard').count()===3);
// 業務内容: Enter連続入力・番号
await p.fill('#fName','瀧澤 瑚乙'); await p.fill('.tk','業務1'); for(let i=2;i<=11;i++){await p.keyboard.press('Enter');await p.keyboard.type('業務'+i);} await p.waitForTimeout(400);
ok('11行入力', await p.locator('#taskBox .tkrow').count()===11);
ok('件数表示', (await p.locator('#taskCount').textContent())==='全11件');
ok('スクロール化', await p.locator('#taskBox').evaluate(e=>e.classList.contains('scroll')));
// 候補
await p.fill('#fGood','新規取引先の与信申請を楽楽販売で入力し、支払条件と与信枠の根拠を整理して説明した。'); await p.waitForTimeout(900);
const cand=await p.locator('#evSuggest .sugrow').last().locator('.sug').allTextContents();
ok('候補に与信関連', cand.some(c=>c.includes('与信')));
// 貼付枠でJSON取り込み
ok('深掘りが既定で1つ開いている', await p.locator('#topicBox .topic').count()===1 && await p.locator('#topicBox .topic .body').first().isVisible());
ok('使い方ガイドがある', await p.locator('.guide li').count()===3);
ok('実績セクションは既定で畳まれている', await p.locator('#evSection.closed').count()===1);
await p.click('#evHead'); await p.waitForTimeout(300);
ok('ヘッダーで開く', await p.locator('#evSection.closed').count()===0);
await p.fill('#evPaste','[{"id":"S063","evidence":"与信申請を差し戻しゼロで通した。"}]');
await p.click('#btnIdCall'); await p.waitForTimeout(400);
ok('JSON取り込み', await p.locator('#evBox .evrow').count()===1);
// モゲジョ: 業務内容から生成・折りたたみ
await p.locator('.mgjbtn').first().click(); await p.waitForTimeout(400);
ok('🔬でモゲジョ生成', (await p.locator('.tp-title').first().inputValue())==='業務1');
await p.locator('.topic .tchev').first().click(); await p.waitForTimeout(200);
ok('折りたたみ', !(await p.locator('.topic .body').first().isVisible()));
await p.locator('.topic .tchev').first().click(); await p.waitForTimeout(200);
await p.locator('.tp-c-none').first().click(); await p.waitForTimeout(200);
ok('制約なしチップ', (await p.locator('.tp-c').first().inputValue()).includes('制約なし'));
await p.locator('.tp-w').first().fill('明日から依頼受領時に工程をNotionへ登録する。');
await p.locator('.tp-j').first().fill('締切のある業務を着手日ベースで先に取る。');
await p.locator('.pickpair button[data-pick="bad"]').first().click(); await p.waitForTimeout(200);
ok('◎△の選択', await p.locator('.pickpair button.on').count()===1);
// 保存＝コピー＋内容保持＋宣言記録
await p.click('#btnSave'); await p.waitForTimeout(700);
ok('保存後も内容が残る', (await p.locator('.tk').first().inputValue())==='業務1');
ok('チャット用にコピー', (await p.evaluate(()=>navigator.clipboard.readText().catch(()=>''))).includes('業務日報'));
ok('宣言を記録', (await p.evaluate(()=>JSON.parse(localStorage.getItem('drSkillMap.v1')).acts.length))===1);
await p.click('#btnSave'); await p.waitForTimeout(600);
ok('再保存で日報が増えない', (await p.locator('#kReports').textContent())==='1');
// スキルマップ
await p.click('.tab[data-v="map"]'); await p.waitForTimeout(400);
ok('ID表示', (await p.locator('#skList .sid').first().textContent())==='S001');
await p.locator('#skList .sk [data-pin]').first().click(); await p.waitForTimeout(300);
ok('ピン止め', (await p.evaluate(()=>JSON.parse(localStorage.getItem('drSkillMap.v1')).pins.length))===1);
await p.selectOption('#qSk','1'); await p.waitForTimeout(300);
ok('スキル種別フィルタ', (await p.locator('#mapCount').textContent()).startsWith('39'));
await p.selectOption('#qSk','');
// スキル表CSV往復
const d1=p.waitForEvent('download'); await p.click('#btnDlSkills2'); const dl1=await d1;
const csv=fs.readFileSync(await dl1.path(),'utf8');
ok('CSV 78行', csv.trim().split('\r\n').length===79);
fs.writeFileSync('/tmp/sk.csv', csv.trim()+'\r\n'+['','レベル2（ミドル）','3. ヒアリング・提案','価格交渉','値引き対応方針','背景を確認し代替案を出す。','商談ロープレで練習できること','考え方・立ち回り'].map(x=>`"${x}"`).join(','));
await p.click('.tab[data-v="data"]'); await p.setInputFiles('#skillFile','/tmp/sk.csv'); await p.waitForTimeout(800);
ok('CSV取り込みで79項目', (await p.evaluate(()=>SKILLS.length))===79);
await p.click('#btnResetSkills'); await p.waitForTimeout(700);
ok('初期表に戻る', (await p.evaluate(()=>SKILLS.length))===78);
// AIモーダル（エージェントモード）
await p.click('.tab[data-v="today"]'); await p.waitForTimeout(300);
if(!(await p.locator('.tk').first().inputValue())) await p.fill('.tk','業務1');
await p.click('#btnAi'); await p.waitForTimeout(600);
ok('AIモーダルが開く', await p.locator('#aiModal.on').count()===1);
const body=await p.locator('#aiPrompt').textContent();
ok('本文のみ（スキル表なし）', !body.includes('S001') && body.includes('本日の業務内容'));
await p.check('#aiSelf'); await p.waitForTimeout(300);
ok('自己完結モードで長文', (await p.locator('#aiPrompt').textContent()).includes('S001'));
await p.uncheck('#aiSelf'); await p.click('#aiModal [data-close]');
// 履歴・評価面談
await p.click('.tab[data-v="hist"]'); await p.waitForTimeout(400);
ok('履歴1件', await p.locator('#histList .hitem').count()===1);
await p.click('.tab[data-v="review"]'); await p.waitForTimeout(400);
ok('実践率KPI', (await p.locator('#revKpi').textContent()).includes('宣言の実践率'));
// 下書き復元
await p.click('.tab[data-v="today"]'); await p.click('#btnNew'); await p.waitForTimeout(400);
await p.fill('.tk','下書きテスト'); await p.waitForTimeout(1200);
await p.reload(); await p.waitForTimeout(700);
ok('下書き復元', (await p.locator('.tk').first().inputValue())==='下書きテスト');
console.log(errs.length?('ERRORS: '+errs.join(' | ')):'ERRORS: none');
await b.close();
