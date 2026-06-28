const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,Header,Footer,
AlignmentType,LevelFormat,HeadingLevel,BorderStyle,WidthType,ShadingType,VerticalAlign,
PageNumber,PageBreak,TableOfContents,TabStopType,TabStopPosition}=require('docx');

const F="Arial";
const NAVY="1F3864", BLUE="2E75B6", LT="D9E2F3", GREY="F2F2F2", GREEN="548235", RED="C00000";

const h1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});
const h2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});
const p=(t,o={})=>new Paragraph({spacing:{after:120,line:288},children:[new TextRun({text:t,...o})]});
const bul=t=>new Paragraph({numbering:{reference:"b",level:0},spacing:{after:60,line:276},children:[new TextRun(t)]});

const border={style:BorderStyle.SINGLE,size:1,color:"BFBFBF"};
const borders={top:border,bottom:border,left:border,right:border};
function cell(text,{w,fill,bold,color,align,size}={}){
  const runs=Array.isArray(text)?text:[new TextRun({text:String(text),bold:!!bold,color:color||"000000",size:size||20})];
  return new TableCell({borders,width:{size:w,type:WidthType.DXA},
    shading:fill?{fill,type:ShadingType.CLEAR}:undefined,
    margins:{top:60,bottom:60,left:100,right:100},verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({alignment:align||AlignmentType.LEFT,children:runs})]});
}
function hrow(cells,widths){return new TableRow({tableHeader:true,children:cells.map((c,i)=>cell(c,{w:widths[i],fill:NAVY,bold:true,color:"FFFFFF",align:i===0?AlignmentType.LEFT:AlignmentType.CENTER}))});}
function row(cells,widths,opt={}){return new TableRow({children:cells.map((c,i)=>cell(c,{w:widths[i],fill:opt.fill,bold:opt.bold,color:opt.colors&&opt.colors[i],align:i===0?AlignmentType.LEFT:(opt.lalign?AlignmentType.LEFT:AlignmentType.RIGHT)}))});}
function tbl(widths,rows){return new Table({width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,rows});}

const styles={
  default:{document:{run:{font:F,size:21}}},
  paragraphStyles:[
    {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:28,bold:true,font:F,color:NAVY},
      paragraph:{spacing:{before:300,after:160},outlineLevel:0,
        border:{bottom:{style:BorderStyle.SINGLE,size:8,color:BLUE,space:4}}}},
    {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:23,bold:true,font:F,color:BLUE},
      paragraph:{spacing:{before:200,after:100},outlineLevel:1}},
  ]
};
const numbering={config:[{reference:"b",levels:[{level:0,format:LevelFormat.BULLET,text:"・",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:460,hanging:240}}}}]}]};
const PAGE={size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}};

// ===== COVER =====
const cover=[
  new Paragraph({spacing:{before:1400}}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:"事 業 計 画 書",bold:true,size:64,color:NAVY,font:F})]}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:1000},children:[new TextRun({text:"美容室 monet（モネ）岡山店　新規開業",bold:true,size:30,color:BLUE,font:F})]}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:"フランチャイズ加盟による美容室出店",size:24})]}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:1400},children:[new TextRun({text:"事業説明 ／ 資金調達計画 ／ 収支計画 ／ FC実績エビデンス",size:22,color:"595959"})]}),
];
const ci=(k,v)=>new TableRow({children:[cell(k,{w:2800,fill:LT,bold:true}),cell(v,{w:5000})]});
const coverTbl=new Table({alignment:AlignmentType.CENTER,width:{size:7800,type:WidthType.DXA},columnWidths:[2800,5000],rows:[
  ci("申請者","島田 和也（個人事業主）"),
  ci("フランチャイザー","株式会社グランフュージョン（monet FC本部）"),
  ci("事業形態","美容室（フランチャイズ加盟）"),
  ci("ブランド","hair salon monet（モネ）"),
  ci("出店地","岡山県岡山市（予定）"),
  ci("開業予定","2026年9月"),
  ci("必要資金","29,462,800円"),
  ci("資金調達","日本政策金融公庫／信用金庫 借入"),
  ci("作成日","2026年6月23日"),
]});
cover.push(coverTbl);
cover.push(new Paragraph({spacing:{before:1600},alignment:AlignmentType.CENTER,children:[new TextRun({text:"本書は融資審査用の資料です。数値は添付Excel「収支計画表」と完全に整合しています。",size:18,color:"808080"})]}));
cover.push(new Paragraph({children:[new PageBreak()]}));

const toc=[h1("目次"),new TableOfContents("Table of Contents",{hyperlink:true,headingStyleRange:"1-2"}),new Paragraph({children:[new PageBreak()]})];

const body=[];
// 1
body.push(h1("1. エグゼクティブサマリー"));
body.push(p("本計画は、実績あるフランチャイズ（FC）「hair salon monet」に加盟し、岡山市に美容室を新規開業するものである。monetは40・50代のエイジング層に特化した高単価コンセプトサロンであり、メニューを「髪質改善」「脳呼吸ヘッドスパ」に絞り込むことで、運営の簡素化・高単価・高リピートを同時に実現する再現性の高いビジネスモデルを確立している。"));
body.push(p("FC本部は既に堀江本店・福岡姪浜院・広島楽々園院など複数店舗を成功裏に運営し、新規出店店舗も開業2ヶ月目で月商200万円超に到達する立ち上がりの速さを実証している。本計画の売上前提は、FC既存店の実績（成熟店平均 月商約266万円）を下回る保守的な水準（1年目末 月商235万円、2年目 月平均250万円）に設定しており、達成確度は高い。"));
body.push(p("開業初年度から営業利益 約121万円・経常利益 約50万円のダブル黒字を確保する。簡易資金繰りベースでも全期間を通じて累計現金残高はプラスを維持する。2年目以降は営業利益約392万円～495万円と安定的に拡大し、借入金（元利均等10年）の返済を十分にカバーする。"));
body.push(new Paragraph({spacing:{before:120,after:120},children:[new TextRun({text:"主要指標（サマリー）",bold:true,size:22,color:NAVY})]}));
body.push(tbl([3000,1600,1600,1600,1560],[
  hrow(["指標","1年目","2年目","3年目","5年目"],[3000,1600,1600,1600,1560]),
  row(["純売上高（円）","25,300,000","30,000,000","31,200,000","32,400,000"],[3000,1600,1600,1600,1560]),
  row(["営業利益（円）","1,211,086","3,922,586","4,504,586","4,948,586"],[3000,1600,1600,1600,1560],{colors:[null,GREEN,GREEN,GREEN,GREEN]}),
  row(["経常利益（円）","504,474","3,282,390","3,932,485","4,517,886"],[3000,1600,1600,1600,1560],{colors:[null,GREEN,GREEN,GREEN,GREEN]}),
  row(["累計現金残高（円）","4,529,028","6,690,680","9,239,305","14,739,387"],[3000,1600,1600,1600,1560],{bold:true}),
]));
body.push(p("（注）添付Excel「収支計画表」年次収支シートより。営業利益率は2年目以降 約12～14%で推移。",{size:18,color:"808080"}));

// 2
body.push(h1("2. 事業概要"));
body.push(h2("2.1 事業者・事業内容"));
body.push(p("申請者は島田和也（個人事業主）。本事業は、株式会社グランフュージョンがフランチャイザー（FC本部）を務める美容室FC「monet」に加盟し、島田がフランチャイジーとして岡山市内に美容室1店舗を出店・運営するものである。主たる収益は髪質改善・脳呼吸ヘッドスパを中心とした技術売上であり、FCブランド・集客ノウハウ・教育システムを活用して早期の経営安定を図る。"));
body.push(h2("2.2 代表者プロフィール（島田 和也）"));
body.push(p("代表者の島田和也は、美容師歴27年のベテランであり、現在もMICHI GROUPの代表（役員）として現場のサロンワークに加え、教育・採用・ブランディング・SNS戦略・新規事業・店舗運営まで経営全般を担っている。プレイヤーとしての高い実績と、経営者としてのマネジメント経験を併せ持つ点が本事業の最大の強みである。"));
body.push(tbl([3600,5560],[
  hrow(["項目","内容"],[3600,5560]),
  row(["美容師歴","27年"],[3600,5560]),
  row(["月間個人売上 最高","10,470,000円（約1,047万円）"],[3600,5560],{colors:[null,GREEN]}),
  row(["客単価","約11,000円"],[3600,5560]),
  row(["指名率","100%"],[3600,5560],{colors:[null,GREEN]}),
  row(["顧客リピート率","94%"],[3600,5560],{colors:[null,GREEN]}),
  row(["現在の役職","MICHI GROUP 代表（役員）"],[3600,5560]),
  row(["受賞","カミカリスマ2026 ノミネート"],[3600,5560]),
]));
body.push(p("「ブランド力 × 現場力 × 経営経験」を強みに、安定した事業運営を目指す。MICHI GROUPとは良好な信頼関係を維持しており、本件は同社を離れるものではなく、27年の経験を自ら責任者として形にする挑戦である。『年齢を重ねても安心して通える・働ける美容室』という想いとmonetの理念が一致し、FCとして岡山での展開を決意した。",{size:18,color:"808080"}));
body.push(h2("2.3 開業の動機・強み"));
body.push(bul("実績あるFCへの加盟：既存店が複数地域で安定稼働し、新規店の立ち上がりパターンが定量的に確立されている。"));
body.push(bul("再現性の高いビジネスモデル：FC本部の集客・採用・店舗運営の仕組みにより、未経験要素のリスクを大幅に低減。"));
body.push(bul("保守的な計画前提：本計画はFC実績の平均を下回る売上前提で組成し、下振れ耐性を確保。"));
body.push(h2("2.4 ロイヤリティ・FC条件"));
body.push(p("フランチャイザー（株式会社グランフュージョン）に支払うFC加盟金は今回限りの特別価格1,500,000円（税込／繰延資産・5年償却。標準価格は3,300,000円）。ロイヤリティは開業4ヶ月目以降、売上高の5%（特別レート。標準は10%）。本計画ではこれらを費用として全額織り込み済みである。"));

// 3 NEW: FC business model
body.push(h1("3. monet FCビジネスモデル"));
body.push(h2("3.1 市場背景とコンセプト"));
body.push(p("美容業界は「少子高齢化による売上減」「情報過多」「労働環境改善の必要性」という三重苦に直面している。monetはこの構造課題に対し、購買力がありながら専用サロンの少ない40・50代エイジング層をターゲットに定め、「五感を満たす空間」「脳呼吸ヘッドスパ」「エイジング毛専用髪質改善剤」の3本柱で差別化する高単価コンセプトサロンを構築した。"));
body.push(h2("3.2 FCの3つの特徴"));
body.push(p("① 市場ニーズとのマッチング精度が高い",{bold:true}));
body.push(bul("集客：40・50代女性は時間・資金に余裕があり、競合の少ないボリュームゾーン。コンプレックス解消が目的のため高単価・高リピートになりやすい。"));
body.push(bul("求人：40代女性美容師（マンツーマン制・デザイン難度が低い・顧客と年齢が近い）の掘り起こしに成功。安定給・ホワイト労働志向のため人件費負担も抑えやすい。"));
body.push(p("② 事業の簡素化",{bold:true}));
body.push(bul("メニューを髪質改善・脳呼吸ヘッドスパに特化（デザインカラー・ブリーチ・パーマ等はNG）。覚える項目・説明・在庫管理が少なく、マニュアル化が容易で単価が安定する。"));
body.push(p("③ コンセプトサロンの実現",/**/{bold:true}));
body.push(bul("「リラクゼーション」を付加し、40・50代ハイエンド層を獲得。原則リストマーケティング（LINE登録→自動予約配信）で広告の費用対効果を可視化する設計。"));
body.push(h2("3.3 高単価・高集客の指標（FC本部基準）"));
body.push(tbl([4680,4680],[
  hrow(["項目","水準"],[4680,4680]),
  row(["オープニング新規","150～230名／単価 約11,000円"],[4680,4680],{lalign:true}),
  row(["通常時 新規","50～70名／単価 約13,000円"],[4680,4680],{lalign:true}),
  row(["想定広告費","月10～15万円"],[4680,4680],{lalign:true}),
  row(["再来（リピート）率","約70%"],[4680,4680],{lalign:true}),
]));
body.push(h2("3.4 FC本部のサポート体制"));
body.push(bul("monetブランドの利用権、全教育・カウンセリング資料の使用権、社員用会員ページの利用。"));
body.push(bul("初期スタッフ教育を本部が代行（4日間）。集客・求人は本部が一元管理。"));
body.push(bul("MT・個別面談を原則禁止する運営設計により、オープン後の月間管理コストは月1～2時間程度（既存加盟2社実績）。"));
body.push(p("（出典）FC本部「monet FC事業説明」資料（本書末尾に添付）。なお同資料記載の加盟金3,300,000円は標準価格であり、本件は今回限りの特別価格1,500,000円（税込）が適用される。",{size:18,color:"808080"}));

// 4 market
body.push(h1("4. 市場性・立地"));
body.push(p("美容業はリピート需要中心のストック型ビジネスであり、景気変動の影響を受けにくく安定的なキャッシュフローを生みやすい。とりわけmonetが狙う40・50代エイジング層は購買力が高く競合が少ないため、高単価・高リピートを実現しやすい。FC「monet」は大阪・福岡・広島など各地で成熟店が月商250万円超を達成しており、特定商圏に依存しない再現性を示している。"));
body.push(p("岡山市は中四国の中核都市であり、十分な人口規模と美容需要を有する。FC本部の出店基準・商圏分析に基づき立地を選定することで、既存店と同等の集客が見込まれる。"));

// 5 evidence
body.push(h1("5. FC実績エビデンス（勝てるビジネスの根拠）"));
body.push(p("以下はFC本部提供の既存monet各店の月次純売上実績である。新規店の立ち上がりの速さと、成熟店の安定水準の双方を示す。"));
body.push(h2("5.1 新規開業店の立ち上がり"));
body.push(tbl([2400,1740,1740,1740,1740],[
  hrow(["店舗","開業1ヶ月目","2ヶ月目","3ヶ月目","4ヶ月目"],[2400,1740,1740,1740,1740]),
  row(["福島院","772,240","1,899,830","2,276,250","2,690,090"],[2400,1740,1740,1740,1740]),
  row(["高槻院","661,250","1,860,850","2,092,690","2,219,790"],[2400,1740,1740,1740,1740]),
  row(["堀江院 2nd","373,750","2,106,900","2,190,865","2,210,300"],[2400,1740,1740,1740,1740]),
]));
body.push(p("→ いずれの新規店も開業2ヶ月目で月商200万円前後に到達。FCの集客力により早期に売上が立ち上がることを実証している。",{bold:true,color:GREEN}));
body.push(h2("5.2 成熟店の安定水準（直近12ヶ月の月平均）"));
body.push(tbl([4680,2340,2340],[
  hrow(["店舗","月平均純売上","状態"],[4680,2340,2340]),
  row(["福岡姪浜院","約3,071,000円","成熟・安定"],[4680,2340,2340]),
  row(["広島楽々園院","約2,492,000円","成熟・安定"],[4680,2340,2340]),
  row(["堀江院（本店）","約2,197,000円","成熟"],[4680,2340,2340]),
  row(["成熟店プール 平均","約2,656,000円","—"],[4680,2340,2340],{bold:true,fill:LT}),
]));
body.push(h2("5.3 FC本部 収益イメージ（参考）"));
body.push(p("FC本部が示す標準的な収益イメージ（席数4・スタッフ4・月商350万円の場合）は以下のとおりで、月間利益 約99万～109万円を見込む。本計画はこれより保守的な売上前提に基づいている。"));
body.push(tbl([4680,4680],[
  hrow(["項目","金額（月額）"],[4680,4680]),
  row(["売上","3,500,000円"],[4680,4680]),
  row(["支出合計（家賃・広告・人件・材料・ロイヤリティ等）","2,315,000～2,415,000円"],[4680,4680]),
  row(["想定利益","985,000～1,085,000円"],[4680,4680],{bold:true,fill:LT}),
]));
body.push(p("本計画は、これらFC実績・本部基準を下回る月商235～250万円を前提としており、保守的かつ達成可能性の高い計画である。",{bold:true}));

// 6 funds
body.push(h1("6. 開業資金・資金調達計画"));
body.push(h2("6.1 必要資金（設備資金＋運転資金）"));
body.push(tbl([4200,2000,3160],[
  hrow(["項目","金額（円）","摘要"],[4200,2000,3160]),
  row(["① 内装・設備工事","18,000,000","概算御見積書 改定（税別16,363,637円）"],[4200,2000,3160]),
  row(["② デザイン設計費","1,358,500","1,235,000円＋税"],[4200,2000,3160]),
  row(["③ その他諸経費・備品代","1,000,000","什器備品一式"],[4200,2000,3160]),
  row(["④ 給湯機（ボイラー）","1,540,000","御見積書（取付費込・税込）"],[4200,2000,3160]),
  row(["⑤ シャンプー設備（現金購入）","668,800","チェア195千×2＋ボウル109千×2（税込）"],[4200,2000,3160]),
  row(["⑥ FC加盟金","1,500,000","特別価格・税込（標準3,300千円）繰延資産"],[4200,2000,3160]),
  row(["⑦ 物件取得費","1,150,000","敷金600千・礼金220千・仲介220千・保証110千（あおい不動産請求書）"],[4200,2000,3160]),
  row(["設備資金 計","25,217,300",""],[4200,2000,3160],{bold:true,fill:GREY}),
  row(["⑧ 運転資金","4,245,500","開業後約2ヶ月分＋販促・予備費"],[4200,2000,3160]),
  row(["開業資金 総額","29,462,800",""],[4200,2000,3160],{bold:true,fill:LT}),
]));
body.push(p("（注）シャンプーチェア2台・シャンプーボウル2台は現金一括購入し設備資金に計上。別途マッサージ付シャンプーチェア2台（1台110万円税込）はリース契約（サムシング御見積／月額59,400円・税込）とし費用計上。火災保険（タフビズ事業活動総合保険・設備什器500万円補償）は年額37,880円を損害保険料として計上している。",{size:18,color:"808080"}));
body.push(h2("6.2 資金調達"));
body.push(tbl([4680,2340,2340],[
  hrow(["調達区分","金額（円）","備考"],[4680,2340,2340]),
  row(["自己資金（充当額）","0","保有10,000,000円は今回使用せず"],[4680,2340,2340]),
  row(["借入金（公庫・信用金庫）","29,462,800","全額借入で試算"],[4680,2340,2340]),
  row(["調達 計","29,462,800","必要資金と一致"],[4680,2340,2340],{bold:true,fill:LT}),
  row(["（参考）保有自己資金","10,000,000","返済・運転の予備原資として確保"],[4680,2340,2340],{fill:LT}),
]));
body.push(p("本計画は審査上保守的に全額借入で試算しているが、申請者は別途10,000,000円の自己資金を保有している（本計画では使用しない）。これは返済・運転の十分な予備原資であり、財務余力の高さを示す。",{size:18,color:"808080"}));

// 7 P&L
body.push(h1("7. 収支計画（5ヶ年）"));
body.push(p("売上原価率12%、人件費・家賃等の固定費、ロイヤリティ（4ヶ月目以降5%・特別レート）、損害保険料、減価償却（設備10年）、加盟金償却（5年）を織り込んだ年次損益は以下のとおり。"));
body.push(tbl([2760,1320,1320,1320,1320,1320],[
  hrow(["勘定科目","1年目","2年目","3年目","4年目","5年目"],[2760,1320,1320,1320,1320,1320]),
  row(["純売上高","25,300","30,000","31,200","31,800","32,400"],[2760,1320,1320,1320,1320,1320]),
  row(["売上総利益","22,264","26,400","27,456","27,984","28,512"],[2760,1320,1320,1320,1320,1320]),
  row(["販売管理費 計","21,053","22,477","22,951","23,257","23,563"],[2760,1320,1320,1320,1320,1320]),
  row(["営業利益","1,211","3,923","4,505","4,727","4,949"],[2760,1320,1320,1320,1320,1320],{bold:true}),
  row(["経常利益","504","3,282","3,932","4,224","4,518"],[2760,1320,1320,1320,1320,1320],{bold:true}),
  row(["税引後利益","353","2,298","2,753","2,957","3,163"],[2760,1320,1320,1320,1320,1320]),
]));
body.push(p("単位：千円。詳細は添付Excel「収支計画表」を参照。1年目から営業利益・経常利益ともに黒字を確保し、2年目以降は安定的に拡大する。",{size:18,color:"808080"}));

// 8 repayment
body.push(h1("8. 借入返済計画"));
body.push(p("借入条件：元金29,462,800円、想定年利2.5%（日本政策金融公庫 創業融資想定）、返済期間10年・元利均等返済。"));
body.push(tbl([3120,3120,3120],[
  hrow(["項目","内容","—"],[3120,3120,3120]),
  row(["毎月返済額","277,746円","元利均等"],[3120,3120,3120]),
  row(["年間返済額","3,332,946円","—"],[3120,3120,3120]),
  row(["総返済額（10年）","33,329,463円","—"],[3120,3120,3120]),
  row(["うち利息総額","3,866,663円","—"],[3120,3120,3120]),
]));
body.push(p("返済原資（税引後利益＋減価償却費＋加盟金償却費）は2年目以降 約482万円～570万円規模となり、年間返済額333万円を大きく上回る。1年目も営業黒字と開業時運転資金により累計現金残高はプラスを維持し、返済に支障はない。さらに保有自己資金10,000,000円が予備原資として控えている。",{bold:true}));

// 9 risk
body.push(h1("9. リスクと対応策"));
body.push(bul("売上下振れ：FC実績平均を下回る保守的前提で計画。さらに本部の集客支援・リストマーケティングにより補完。"));
body.push(bul("人材確保：FC本部の採用・教育システムと40代女性美容師の掘り起こしノウハウを活用し、開業時から品質と人員を確保。"));
body.push(bul("初期費用負担：マッサージ付シャンプーチェアをリース化し初期投資を圧縮。運転資金を約2ヶ月分確保し資金繰りに余裕。"));
body.push(bul("金利変動：固定的な返済計画で試算。自己資金充当により借入圧縮も選択可能。"));

// 10 conclusion
body.push(h1("10. 結論"));
body.push(p("本事業は、美容師歴27年・個人売上最高月1,047万円・指名率100%・顧客リピート率94%という卓越した実績を持つ代表者が、MICHI GROUPでの経営参画で培ったマネジメント力と、全国で支持されるmonetブランドのノウハウを掛け合わせて運営する、再現性と勝算の高い事業である。"));
body.push(p("本計画は、実績あるFCの定量データと確立されたビジネスモデルに裏付けられた再現性の高い事業である。売上前提はFC既存店の実績を下回る保守的水準に設定し、2年目以降は安定的な黒字と十分な返済原資を確保する。初年度も資金繰り上の累計現金残高はプラスを維持する。以上より、本事業は融資対象として十分な返済能力と事業性を備えた『勝てるビジネス』であると判断する。"));
body.push(p("※ 本書末尾にFC本部「monet FC事業説明」資料を添付する。",{size:18,color:"808080"}));

const footer=new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
  new TextRun({text:"美容室 monet 岡山店 事業計画書　／　申請者：島田 和也　　",size:16,color:"808080"}),
  new TextRun({text:"- ",size:16,color:"808080"}),new TextRun({children:[PageNumber.CURRENT],size:16,color:"808080"}),new TextRun({text:" -",size:16,color:"808080"})
]})]});

const doc=new Document({styles,numbering,features:{updateFields:true},sections:[
  {properties:{page:PAGE},children:cover},
  {properties:{page:PAGE},footers:{default:footer},children:[...toc,...body]},
]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("monet岡山店_事業計画書.docx",b);console.log("docx saved",b.length);});
