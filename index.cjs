<!-- index.html -->
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8" />
  <title>海外旅行保険 申込フォーム</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    :root { --bg:#f7f7f8; --card:#fff; --text:#111; --muted:#666; --border:#ddd; --brand:#2563eb; --danger:#b91c1c; }
    body { margin:0; font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Noto Sans JP", "Meiryo", sans-serif; background:var(--bg); color:var(--text); }
    .wrap { max-width: 920px; margin: 28px auto 64px; padding: 0 14px; }
    .card { background:var(--card); border:1px solid var(--border); border-radius: 14px; box-shadow: 0 4px 18px rgba(0,0,0,.06); padding: 20px; }
    h1 { margin: 0 0 4px; font-size: 1.6rem; }
    p.lead { margin: 0 0 18px; color: var(--muted); }
    .grid { display: grid; grid-template-columns: 1fr; gap: 14px; }
    @media (min-width: 820px){ .grid-2 { display:grid; grid-template-columns: 1fr 1fr; gap:14px; } }
    label { display:block; font-weight:600; margin: 6px 0; }
    input, select, textarea { width:100%; box-sizing:border-box; padding:10px 12px; border-radius:10px; border:1px solid var(--border); background:#fff; font-size:14px; }
    textarea { min高さ: 90px; resize: vertical; }
    .row { margin-top: 18px; }
    .subhead { margin: 24px 0 8px; font-size: 1.1rem; font-weight: 800; }
    .muted { color: var(--muted); font-size: 12px; }
    .actions { display:flex; gap:12px; margin-top: 20px; }
    button { padding: 12px 18px; border:0; border-radius:10px; font-weight:700; cursor:pointer; }
    button.primary { background: var(--brand); color: #fff; }
    button.secondary { background: #e5e7eb; }
    .error { color: var(--danger); margin-top: 8px; }
    fieldset { border:1px dashed var(--border); border-radius: 12px; padding: 12px; }
    legend { font-weight: 700; font-size: 0.95rem; color:#333; }
    .disabled { opacity: .6; }
    .inline { display:flex; gap:10px; align-items:center; }
    .hint { font-size:12px; color:var(--muted); margin-top:4px;}
    table { width:100%; border-collapse: collapse; margin-top:10px; }
    th, td { border:1px solid var(--border); padding:8px; vertical-align: top; }
    th { background:#f3f4f6; text-align:left; }
    .smallmuted { font-size:12px; color:var(--muted); margin-top:6px; }
    .kicker { font-weight:700; margin-top:6px; }
    .pill { display:inline-block; padding:2px 8px; border-radius:999px; border:1px solid var(--border); font-size:12px; color:#444; }
    .radio-group label { display:inline-block; margin-right:14px; font-weight:normal; }
    .two { display:grid; grid-template-columns: 1fr 1fr; gap: 10px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>海外旅行保険 申込フォーム</h1>
      <p class="lead">年齢・持病・加入期間によって選択できるプランが自動的に変わります。</p>

      <form id="appForm">
        <!-- ===== 上段：プラン判定（ロジックはそのまま） ===== -->
        <div class="grid">
          <div>
            <label>お客様ID</label>
            <input type="text" id="CustomerID" name="CustomerID" required />
          </div>
          <div>
            <label>就学予定の学校名</label>
            <select id="SchoolName" name="SchoolName" required>
              <option value="">選択してください</option>
            </select>
          </div>

          <div>
            <label>年齢</label>
            <input type="number" id="Age" name="Age" min="0" max="120" required />
          </div>

          <!-- 病気 検索＋セレクト -->
          <div>
            <label>持病</label>
            <div class="inline">
              <input id="IllnessSearch" type="text" placeholder="持病を検索…" autocomplete="off" />
              <select id="Illness" name="Illness" disabled>
                <option value="" selected>該当なし</option>
              </select>
            </div>
            <div class="hint"><span id="illnessCount">0</span> 件一致（検索は上のテキストボックスで）</div>
          </div>

          <div>
            <label>加入期間</label>
            <select id="CoverageLength" name="CoverageLength" required>
              <option value="" selected>選択してください</option>
            </select>
            <div class="muted">プランの適用可否・料金計算に使用します。</div>
          </div>

          <div>
            <label>危険な活動の予定</label>
            <select id="DangerousActivities" name="DangerousActivities">
              <option value="No" selected>いいえ</option>
              <option value="Yes">はい</option>
            </select>
          </div>

          <div>
            <label>希望プラン</label>
            <select id="Plan" name="Plan" required disabled>
              <option value="" selected>年齢・病気・期間を選択してください</option>
            </select>
            <div class="muted" id="planMeta"></div>
            <!-- ★ ミラー（hidden） -->
            <input type="hidden" name="Plan" id="PlanHidden" />
          </div>

          <div>
            <label>開始日</label>
            <input type="date" id="StartDate" name="CoverageStart" required disabled />
            <!-- ★ ミラー（hidden） -->
            <input type="hidden" name="CoverageStart" id="CoverageStartHidden" />
          </div>

          <div>
            <label>終了日（自動計算）</label>
            <input type="date" id="EndDate" name="CoverageEnd" readonly />
            <!-- ★ 念のためミラー（readonlyでも送られるが保険で） -->
            <input type="hidden" name="CoverageEnd" id="CoverageEndHidden" />
          </div>

          <div>
            <label>料金（自動）</label>
            <input type="text" id="Price" name="Price" readonly />
            <!-- ★ 念のためミラー -->
            <input type="hidden" name="Price" id="PriceHidden" />
          </div>
        </div>

        <div class="row muted">
          ステータス: <span id="eligStatus" class="pill">初期化中…</span>
          <span id="ruleCounts" class="pill" style="margin-left:8px;"></span>
        </div>

        <!-- 保証表 -->
        <div class="kicker">プランの主な補償内容</div>
        <div class="smallmuted">該当プランの行があれば表示します（なければ分類にフォールバック）。</div>
        <div id="coverageWrap">
          <table id="coverageTable" style="display:none;">
            <thead id="coverageHead"></thead>
            <tbody id="coverageBody"></tbody>
          </table>
          <div id="coverageEmpty" class="muted" style="display:none; margin-top:6px;">該当データがありません。</div>
        </div>

        <!-- ===== 告知事項（補償表の直下） ===== -->
        <div class="subhead">告知事項</div>
        <fieldset>
          <legend>健康状態・渡航先・職務</legend>

          <label>1) 現在、ケガや病気で医師の治療・投薬を受けていますか？</label>
          <div class="radio-group">
            <label><input type="radio" name="Q1_TreatmentNow" value="はい" required> はい</label>
            <label><input type="radio" name="Q1_TreatmentNow" value="いいえ"> いいえ</label>
          </div>
          <div id="Q1_Detail" style="display:none; margin-top:6px;">
            <label>傷病名を記載してください</label>
            <input type="text" name="Q1_DiseaseName" placeholder="例）糖尿病 など" />
          </div>

          <label style="margin-top:8px;">2) これまで継続して１ヶ月以上入院したこと、または脳疾患・心疾患・ガンを患ったことがありますか？</label>
          <div class="radio-group">
            <label><input type="radio" name="Q2_SeriousHistory" value="はい" required> はい</label>
            <label><input type="radio" name="Q2_SeriousHistory" value="いいえ"> いいえ</label>
          </div>
          <div id="Q2_Detail" style="display:none; margin-top:6px;">
            <label>詳細を記載してください</label>
            <textarea name="Q2_DetailText" rows="3"></textarea>
          </div>

          <label style="margin-top:8px;">3) 過去３年間に携行品の保険金を５回以上請求・受領されていますか？</label>
          <div class="radio-group">
            <label><input type="radio" name="Q3_LuggageClaims5Plus" value="はい" required> はい</label>
            <label><input type="radio" name="Q3_LuggageClaims5Plus" value="いいえ"> いいえ</label>
          </div>

          <label style="margin-top:8px;">4) 同一補償内容の他の保険契約（共済を含む）はありますか？</label>
          <div class="radio-group">
            <label><input type="radio" name="Q4_DuplicateContracts" value="はい" required> はい</label>
            <label><input type="radio" name="Q4_DuplicateContracts" value="いいえ"> いいえ</label>
          </div>
          <div id="Q4_Detail" style="display:none; margin-top:6px;">
            <div class="smallmuted">当てはまるものを選択（複数可）。6=その他の場合は自由記載も。</div>
            <div>
              <label><input type="checkbox" name="DupContract_1"> 1. 海外旅行保険</label>
              <label><input type="checkbox" name="DupContract_2"> 2. 普通傷害保険</label>
              <label><input type="checkbox" name="DupContract_3"> 3. 家族傷害保険</label>
              <label><input type="checkbox" name="DupContract_4"> 4. 傷害総合保険</label>
              <label><input type="checkbox" name="DupContract_5"> 5. 交通事故傷害保険</label>
              <label><input type="checkbox" name="DupContract_6"> 6. 共済などその他</label>
            </div>
            <div class="two" style="margin-top:6px;">
              <div>
                <label>その他（自由記載）</label>
                <input type="text" name="DupContract_OtherText" />
              </div>
              <div>
                <label>保険会社名など</label>
                <input type="text" name="DupContract_CompanyText" />
              </div>
            </div>
            <div style="margin-top:6px;">
              <label>傷害死亡保険金額</label>
              <input type="text" name="DeathBenefitAmount" placeholder="例）1,000万円" />
            </div>
          </div>

          <label style="margin-top:8px;">5) 今回の旅行先をお選びください</label>
          <select name="Q5_DestinationRegion" id="Q5_DestinationRegion" required>
            <option value="">選択してください</option>
            <option>アジア</option>
            <option>ヨーロッパ</option>
            <option>オセアニア</option>
            <option>北米</option>
            <option>中南米</option>
            <option>アフリカ</option>
            <option>中東</option>
            <option>その他</option>
          </select>
          <div id="Q5_OtherWrap" style="display:none; margin-top:6px;">
            <label>その他（自由記載）</label>
            <input type="text" name="DestinationOtherText" />
          </div>

          <label style="margin-top:8px;">6) 次の国（イラン、スーダン、シリア、クリミア地域、キューバ）が含まれますか？</label>
          <div class="radio-group">
            <label><input type="radio" name="Q6_SanctionedCountries" value="はい" required> はい</label>
            <label><input type="radio" name="Q6_SanctionedCountries" value="いいえ"> いいえ</label>
          </div>

          <label style="margin-top:8px;">7) ご旅行中に従事する職務はありますか？（ワーホリ、Co-op留学は「いいえ」）</label>
          <div class="radio-group">
            <label><input type="radio" name="Q7_JobDuringTravel" value="はい" required> はい</label>
            <label><input type="radio" name="Q7_JobDuringTravel" value="いいえ"> いいえ</label>
          </div>

          <div id="Q7_Detail" style="display:none; margin-top:6px;">
            <label>詳細（任意）</label>
            <textarea name="WorkDuringTravelText" rows="2" placeholder="任意"></textarea>
          </div>
        </fieldset>

        <!-- ===== 氏名（告知事項の後に配置） ===== -->
        <div class="subhead">氏名</div>
        <fieldset>
          <legend>申込人（漢字・カナ）</legend>
          <div class="two">
            <div>
              <label>苗字（漢字）</label>
              <input type="text" id="ApplicantLastKanji" name="ApplicantLastKanji" />
            </div>
            <div>
              <label>名前（漢字）</label>
              <input type="text" id="ApplicantFirstKanji" name="ApplicantFirstKanji" />
            </div>
          </div>
          <div class="two">
            <div>
              <label>苗字（カナ）</label>
              <input type="text" id="ApplicantLastKana" name="ApplicantLastKana" />
            </div>
            <div>
              <label>名前（カナ）</label>
              <input type="text" id="ApplicantFirstKana" name="ApplicantFirstKana" />
            </div>
          </div>
        </fieldset>

        <!-- ===== 申込人／旅行者 情報 ===== -->
        <div class="subhead">申込人・旅行者情報</div>

        <fieldset>
          <legend>申込人</legend>
          <div>
            <label>郵便番号（7桁）</label>
            <input type="text" id="ApplicantPostal" name="ApplicantPostal" placeholder="例）0010001" />
          </div>
          <div>
            <label>住所</label>
            <input type="text" id="ApplicantAddress" name="ApplicantAddress" />
          </div>
          <div class="two">
            <div>
              <label>電話番号</label>
              <input type="text" id="ApplicantPhone" name="ApplicantPhone" placeholder="09012345678 など" />
            </div>
            <div>
              <label>電話の種別</label>
              <select id="PhoneType" name="PhoneType">
                <option value="携帯" selected>携帯</option>
                <option value="自宅">自宅</option>
              </select>
            </div>
          </div>
          <div class="two">
            <div>
              <label>生年月日（申込人）</label>
              <input type="date" id="ApplicantBirth" name="ApplicantBirth" />
            </div>
            <div class="muted" style="display:flex; align-items:flex-end;">※PDFでは元号A/B＋YY/MM/DDで出力</div>
          </div>
        </fieldset>

        <fieldset>
          <legend>旅行者</legend>
          <div class="two">
            <div>
              <label>姓（ローマ字）</label>
              <input type="text" id="TravelerLastEn" name="TravelerLastEn" placeholder="YAMADA" />
            </div>
            <div>
              <label>名（ローマ字）</label>
              <input type="text" id="TravelerFirstEn" name="TravelerFirstEn" placeholder="TARO" />
            </div>
          </div>
          <div>
            <label>性別</label>
            <select id="TravelerSex" name="TravelerSex">
              <option value="">選択してください</option>
              <option value="男性">男性</option>
              <option value="女性">女性</option>
            </select>
          </div>

          <div style="margin-top:6px;">
            <label><input type="checkbox" id="SameAsTraveler" name="SameAsTraveler" /> 申込人と旅行者は同一です（申込人の住所・電話・生年月日を使用）</label>
          </div>

          <div id="TravelerDetailsSection">
            <div>
              <label>生年月日（旅行者）</label>
              <input type="date" id="TravelerBirth" name="TravelerBirth" />
            </div>
            <div>
              <label>郵便番号（7桁）</label>
              <input type="text" id="TravelerPostal" name="TravelerPostal" placeholder="例）0010001" />
            </div>
            <div>
              <label>住所（旅行者）</label>
              <input type="text" id="TravelerAddress" name="TravelerAddress" />
            </div>
            <div>
              <label>電話番号（旅行者）</label>
              <input type="text" id="TravelerPhone" name="TravelerPhone" placeholder="09012345678 など" />
            </div>
            <div>
              <label>電話の種別（旅行者）</label>
              <select id="TravelerPhoneType" name="TravelerPhoneType">
                <option value="携帯" selected>携帯</option>
                <option value="自宅">自宅</option>
                <option value="勤務先">勤務先</option>
              </select>
            </div>
          </div>
        </fieldset>

        <!-- ===== 郵送先住所 ===== -->
        <fieldset>
          <legend>郵送先住所</legend>
          <label>ご帰国後に郵送物をお受け取りになれるご住所</label>
          <textarea name="AddressAfterReturn" rows="3"></textarea>

          <label>ご出発前に郵便物をお受け取りになれるご住所</label>
          <textarea name="AddressBeforeDeparture" rows="3"></textarea>
        </fieldset>

        <!-- 任意メモ -->
        <div class="grid">
          <div>
            <label>備考</label>
            <textarea id="Notes" name="Notes" placeholder="任意メモ"></textarea>
          </div>
        </div>

        <!-- hidden: 申込日（当日を送信時にセット） -->
        <input type="hidden" id="ApplyDate" name="ApplyDate" />

        <div id="errorBox" class="error" style="display:none;"></div>
        <div class="actions">
          <button type="submit" id="submitBtn" class="primary">PDFを作成</button>
          <button type="reset" class="secondary" id="resetBtn">リセット</button>
        </div>
        <div class="muted" style="margin-top:8px;">
          送信するとPDFフィラーにデータを送信し、Driveに保存されたPDFのリンクを表示します。
        </div>
      </form>
    </div>
  </div>

  <script>
    // ===== In-memory rules =====
    let RULES = { plans: [], illnesses: [], coverages: [], durationHeaders: [] };
    let ILLNESS_ALL = [];

    // ===== Elements =====
    const el = (id)=>document.getElementById(id);
    const f = {
      CustomerID: el('CustomerID'),
      SchoolName: el('SchoolName'),
      Age: el('Age'),
      IllnessSearch: el('IllnessSearch'),
      Illness: el('Illness'),
      Dangerous: el('DangerousActivities'),
      CoverageLength: el('CoverageLength'),
      eligStatus: el('eligStatus'),
      ruleCounts: el('ruleCounts'),
      Plan: el('Plan'),
      PlanMeta: el('planMeta'),
      Start: el('StartDate'),
      End: el('EndDate'),
      Price: el('Price'),
      coverageHead: el('coverageHead'),
      coverageBody: el('coverageBody'),
      coverageTable: el('coverageTable'),
      coverageEmpty: el('coverageEmpty'),

      // Applicant / Traveler
      ApplicantLastKanji: el('ApplicantLastKanji'),
      ApplicantFirstKanji: el('ApplicantFirstKanji'),
      ApplicantLastKana: el('ApplicantLastKana'),
      ApplicantFirstKana: el('ApplicantFirstKana'),
      ApplicantPostal: el('ApplicantPostal'),
      ApplicantAddress: el('ApplicantAddress'),
      ApplicantPhone: el('ApplicantPhone'),
      ApplicantBirth: el('ApplicantBirth'),
      PhoneType: el('PhoneType'),

      TravelerLastEn: el('TravelerLastEn'),
      TravelerFirstEn: el('TravelerFirstEn'),
      TravelerSex: el('TravelerSex'),
      TravelerBirth: el('TravelerBirth'),
      TravelerPostal: el('TravelerPostal'),
      TravelerAddress: el('TravelerAddress'),
      TravelerPhone: el('TravelerPhone'),
      TravelerPhoneType: el('TravelerPhoneType'),
      SameAsTraveler: el('SameAsTraveler'),

      ApplyDate: el('ApplyDate'),

      //告知 動的表示
      DestRegion: el('Q5_DestinationRegion'),
    };
    const errorBox = el('errorBox');
    const submitBtn = el('submitBtn');

    function setLoading(loading){
      if (loading){ submitBtn.textContent='作成中…'; submitBtn.disabled=true; }
      else { submitBtn.textContent='PDFを作成'; submitBtn.disabled=false; }
    }
    function showError(m){ errorBox.style.display='block'; errorBox.textContent=m; }
    function clearError(){ errorBox.style.display='none'; errorBox.textContent=''; }

    // ===== Duration headers =====
    function parseDurationHeader(header){
      const h = String(header).replace(/\u200B/g,'').trim();
      const mDays = h.match(/^(\d+)日.*まで$/);
      if (mDays) return { type:'days', value:Number(mDays[1]), label:`${mDays[1]}日まで`, header };
      const mRange = h.match(/^(\d+).+?(\d+)日.*まで$/);
      if (mRange) return { type:'days', value:Number(mRange[2]), label:`${mRange[1]}〜${mRange[2]}日まで`, header };
      const mMonths = h.match(/^(\d+)(?:ヶ|ケ)?月.*まで$/);
      if (mMonths) return { type:'months', value:Number(mMonths[1]), label:`${mMonths[1]}ヶ月まで`, header };
      return null;
    }
    function buildDurationOptions(){
      const seen = new Map();
      (RULES.durationHeaders||[]).forEach(h=>{
        const info = parseDurationHeader(h);
        if (info) seen.set(info.type+':'+info.value+':'+info.header, info);
      });
      const arr = Array.from(seen.values()).sort((a,b)=>{
        if (a.type!==b.type) return a.type==='days'?-1:1;
        return a.value-b.value;
      });
      f.CoverageLength.innerHTML = '<option value="" selected>選択してください</option>';
      arr.forEach(d=>{
        const opt=document.createElement('option');
        opt.value = JSON.stringify(d);
        opt.textContent = d.label;
        f.CoverageLength.appendChild(opt);
      });
    }

    // ===== School list loading =====
    function loadSchoolList(){
      const schoolSelect = f.SchoolName;
      schoolSelect.innerHTML = '<option value="">読み込み中...</option>';
      
      google.script.run
        .withSuccessHandler(data => {
          schoolSelect.innerHTML = '<option value="">選択してください</option>';
          if (data && data.ok && data.schools && Array.isArray(data.schools)) {
            data.schools.forEach(school => {
              const option = document.createElement('option');
              option.value = school;
              option.textContent = school;
              schoolSelect.appendChild(option);
            });
          } else {
            schoolSelect.innerHTML = '<option value="">学校データの読み込みに失敗しました</option>';
          }
        })
        .withFailureHandler(error => {
          console.error('School data loading error:', error);
          schoolSelect.innerHTML = '<option value="">学校データの読み込みに失敗しました</option>';
        })
        .getSchools();
    }

    // ===== Illness list + search =====
    function normalizeJa(s){ return String(s||'').toLowerCase().replace(/\s+/g,''); }
    function buildIllnessList(){
      const rows = (RULES.illnesses||[]);
      ILLNESS_ALL = [];
      if (!rows.length){ renderIllnessOptions([]); return; }
      const headers = Object.keys(rows[0]||{});
      const col =
        headers.find(h=>/病名|疾患名|疾病名|病気名|illness/i.test(h)) ||
        headers.find(h=>/病|疾患|ill/i.test(h)) || '病名';
      const names = new Set();
      rows.forEach(r=>{
        const name = String(r[col]||'').trim();
        if (name && !names.has(name)){ names.add(name); ILLNESS_ALL.push(name); }
      });
      ILLNESS_ALL.push('なし');
      renderIllnessOptions(ILLNESS_ALL);
      f.Illness.disabled = false;
    }
    function renderIllnessOptions(list){
      const sel=f.Illness;
      sel.innerHTML='';
      const def=document.createElement('option');
      def.value=''; def.textContent='該当なし';
      sel.appendChild(def);
      list.forEach(n=>{
        const o=document.createElement('option'); o.value=n; o.textContent=n; sel.appendChild(o);
      });
      el('illnessCount').textContent=String(list.length);
      const cur=sel.getAttribute('data-current');
      if (cur && list.includes(cur)) sel.value=cur; else sel.value='';
    }
    function filterIllnessOptions(q){
      const s=normalizeJa(q);
      if (!s) return renderIllnessOptions(ILLNESS_ALL);
      renderIllnessOptions(ILLNESS_ALL.filter(n=>normalizeJa(n).includes(s)));
    }

    // ===== Coverage table =====
    function renderCoverageFor(planName, planClass){
      const rows=(RULES.coverages||[]);
      if (!rows.length){ f.coverageTable.style.display='none'; f.coverageEmpty.style.display='block'; f.coverageEmpty.textContent='データなし'; return; }
      let matched=rows.filter(r => (r['プラン名']||r['プラン']||'').trim()===planName);
      if (!matched.length && planClass) matched=rows.filter(r=>(r['プラン分類']||'').trim()===planClass);
      if (!matched.length){ f.coverageTable.style.display='none'; f.coverageEmpty.style.display='block'; f.coverageEmpty.textContent='該当なし'; return; }
      const first=matched[0];
      const keys=Object.keys(first).filter(k=>k && !/^\s*$/.test(k));
      f.coverageHead.innerHTML='<tr>'+keys.map(k=>`<th>${escapeHtml(k)}</th>`).join('')+'</tr>';
      f.coverageBody.innerHTML='';
      matched.forEach(r=>{
        const tr=document.createElement('tr');
        tr.innerHTML=keys.map(k=>`<td>${escapeHtml(r[k]||'')}</td>`).join('');
        f.coverageBody.appendChild(tr);
      });
      f.coverageEmpty.style.display='none'; f.coverageTable.style.display='table';
    }

    // ===== Eligibility =====
    function eligiblePlans(age, illnessInfo, durationDesc, dangerousFlag){
      if (!durationDesc) return [];
      const header=durationDesc.header;
      return (RULES.plans||[]).filter(p=>{
        const ageMin=toNum(p['年齢_Min'],-Infinity);
        const ageMax=toNum(p['年齢_Max'],Infinity);
        if (isFinite(ageMin)&&age<ageMin) return false;
        if (isFinite(ageMax)&&age>ageMax) return false;
        if (illnessInfo && illnessInfo.planClass){
          const pc=(p['プラン分類']||'').trim();
          if (pc!==illnessInfo.planClass) return false;
        }
        const priceCell=(p[header]||'').trim(); if (!priceCell) return false;
        const dangerCol=p['危険活動可'];
        if (dangerCol && dangerousFlag==='Yes' && String(dangerCol).trim().toLowerCase()!=='yes') return false;
        return true;
      });
    }
    function toNum(v, fb=NaN){ if(v==null) return fb; const n=Number(String(v).replace(/[,¥\s]/g,'')); return isNaN(n)?fb:n; }
    function getIllnessInfo(name){
      if (!name || name==='なし') return null;
      const row=(RULES.illnesses||[]).find(r=>{
        const hs=Object.keys(r||{});
        const col=hs.find(h=>/病名|疾患名|疾病名|病気名|illness/i.test(h)) || '病名';
        return String(r[col]||'').trim()===name.trim();
      });
      if (!row) return null;
      return { name, planClass:(row['該当プラン']||'').trim() };
    }

    function updateEligibility(){
      clearError();
      const age=Number(f.Age.value||NaN);
      const illnessInfo=getIllnessInfo(f.Illness.value);
      const durationDesc=f.CoverageLength.value?JSON.parse(f.CoverageLength.value):null;
      const dangerousFlag=f.Dangerous.value;
      if (!isFinite(age)||!durationDesc){ f.eligStatus.textContent='年齢・期間を入力してください。'; lockPlan(true); return; }
      const plans=eligiblePlans(age, illnessInfo, durationDesc, dangerousFlag);
      renderPlanOptions(plans, durationDesc);
      if (!plans.length){ f.eligStatus.textContent='適用可能なプランがありません。'; lockPlan(true); }
      else { const t=illnessInfo&&illnessInfo.planClass?`（既往症→${illnessInfo.planClass}）`:''; f.eligStatus.textContent=`選択可能プラン: ${plans.length} ${t}`; lockPlan(false); }
    }

    function lockPlan(lock){
      f.Plan.disabled=!!lock; f.Start.disabled=!!lock;
      if (lock){ f.Plan.innerHTML='<option value="" selected>年齢・病気・期間を選択してください</option>'; f.Price.value=''; f.End.value=''; f.coverageTable.style.display='none'; f.coverageEmpty.style.display='none'; }
    }

    function renderPlanOptions(plans, durationDesc){
      const header=durationDesc.header;
      f.Plan.innerHTML='<option value="" selected>選択してください</option>';
      plans.forEach(p=>{
        const name=(p['プラン名']||'').trim(); const cls=(p['プラン分類']||'').trim(); const price=(p[header]||'').trim();
        const opt=document.createElement('option');
        opt.value=name; opt.dataset.class=cls; opt.dataset.header=header; opt.dataset.price=price; opt.textContent=name;
        f.Plan.appendChild(opt);
      });
      f.PlanMeta.textContent=''; f.Price.value=''; f.End.value=''; f.coverageTable.style.display='none'; f.coverageEmpty.style.display='none';
    }

    function computeEndDate(){
      const start=f.Start.value; if(!start){ f.End.value=''; return; }
      const opt=f.Plan.options[f.Plan.selectedIndex]; if(!opt) return;
      const d=parseDurationHeader(opt.dataset.header||''); if(!d) return;
      const sd=new Date(start+'T00:00:00'); let ed=new Date(sd);
      if (d.type==='days'){ ed.setDate(ed.getDate()+ (d.value-1)); } else { ed.setMonth(ed.getMonth()+d.value); ed.setDate(ed.getDate()-1); }
      f.End.value = `${ed.getFullYear()}-${String(ed.getMonth()+1).padStart(2,'0')}-${String(ed.getDate()).padStart(2,'0')}`;
    }
    function updatePriceMeta(){
      const opt=f.Plan.options[f.Plan.selectedIndex]; if(!opt) return;
      f.Price.value=opt.dataset.price||''; f.PlanMeta.textContent=opt.dataset.class?`分類: ${opt.dataset.class}`:''; computeEndDate(); renderCoverageFor(opt.value, opt.dataset.class||'');
    }

    // ===== 告知：条件表示 =====
    const Q1 = document.getElementsByName('Q1_TreatmentNow');
    const Q2 = document.getElementsByName('Q2_SeriousHistory');
    const Q4 = document.getElementsByName('Q4_DuplicateContracts');
    const Q8 = document.getElementsByName('Q7_JobDuringTravel'); // 7で8の詳細表示
    function onRadioGroupChange(nodeList, targetId, showIfValue='はい'){
      const val = [...nodeList].find(r=>r.checked)?.value || '';
      document.getElementById(targetId).style.display = (val === showIfValue) ? 'block' : 'none';
    }
    [...Q1].forEach(r=>r.addEventListener('change', ()=>onRadioGroupChange(Q1,'Q1_Detail')));
    [...Q2].forEach(r=>r.addEventListener('change', ()=>onRadioGroupChange(Q2,'Q2_Detail')));
    [...Q4].forEach(r=>r.addEventListener('change', ()=>onRadioGroupChange(Q4,'Q4_Detail')));
    [...Q8].forEach(r=>r.addEventListener('change', ()=>{
      const val = [...Q8].find(r=>r.checked)?.value || '';
      document.getElementById('Q7_Detail').style.display = (val === 'はい') ? 'block' : 'none';
    }));
    f.DestRegion.addEventListener('change', ()=>{
      document.getElementById('Q5_OtherWrap').style.display = (f.DestRegion.value === 'その他') ? 'block' : 'none';
    });

    // ===== 同一人物同期（住所・電話・生年月日） =====
    f.SameAsTraveler.addEventListener('change', handleSameAsTravelerChange);
    ['ApplicantPostal','ApplicantAddress','ApplicantPhone','PhoneType','ApplicantBirth'].forEach(id=>{
      el(id).addEventListener('input', ()=>{ if (f.SameAsTraveler.checked) syncTravelerFromApplicant(); });
      el(id).addEventListener('change', ()=>{ if (f.SameAsTraveler.checked) syncTravelerFromApplicant(); });
    });
    
    function handleSameAsTravelerChange(){
      const isSamePerson = f.SameAsTraveler.checked;
      const travelerDetailsSection = document.getElementById('TravelerDetailsSection');
      const travelerBirthField = document.getElementById('TravelerBirth');
      
      if (isSamePerson) {
        // Hide traveler address and phone fields
        if (travelerDetailsSection) {
          travelerDetailsSection.style.display = 'none';
        }
        
        // Auto-fill traveler data from applicant
        syncTravelerFromApplicant();
        
        // Set traveler birthday = applicant birthday
        if (f.TravelerBirth && f.ApplicantBirth) {
          f.TravelerBirth.value = f.ApplicantBirth.value;
        }
        
        // Disable traveler birthday field (it's auto-filled)
        if (travelerBirthField) {
          travelerBirthField.disabled = true;
          travelerBirthField.style.opacity = '0.6';
        }
        
      } else {
        // Show traveler address and phone fields
        if (travelerDetailsSection) {
          travelerDetailsSection.style.display = 'block';
        }
        
        // Clear traveler data (user needs to fill independently)
        if (f.TravelerPostal) f.TravelerPostal.value = '';
        if (f.TravelerAddress) f.TravelerAddress.value = '';
        if (f.TravelerPhone) f.TravelerPhone.value = '';
        if (f.TravelerPhoneType) f.TravelerPhoneType.value = '携帯'; // Reset to default
        if (f.TravelerBirth) f.TravelerBirth.value = '';
        
        // Enable traveler birthday field
        if (travelerBirthField) {
          travelerBirthField.disabled = false;
          travelerBirthField.style.opacity = '1';
        }
      }
    }
    
    function syncTravelerFromApplicant(){
      const on=f.SameAsTraveler.checked;
      if (on){
        // Only copy birthday when checkbox is checked
        f.TravelerBirth.value      = f.ApplicantBirth.value;
        // Keep postal, address, phone, and phone type blank
        f.TravelerPostal.value     = '';
        f.TravelerAddress.value    = '';
        f.TravelerPhone.value      = '';
        f.TravelerPhoneType.value  = '';
      }
    }

    // ===== ★ 送信直前：disabled/readonly を hidden にミラー =====
    function syncHiddenBeforeSubmit(){
      // Plan → PlanHidden
      const planHidden = document.getElementById('PlanHidden');
      planHidden.value = f.Plan.value || '';

      // CoverageStart → CoverageStartHidden
      const covStartHidden = document.getElementById('CoverageStartHidden');
      covStartHidden.value = f.Start.value || '';

      // CoverageEnd → CoverageEndHidden（readonly だが保険で）
      const covEndHidden = document.getElementById('CoverageEndHidden');
      covEndHidden.value = f.End.value || '';

      // Price → PriceHidden（readonly だが保険で）
      const priceHidden = document.getElementById('PriceHidden');
      priceHidden.value = f.Price.value || '';
    }

    // ===== Submit =====
    document.getElementById('appForm').addEventListener('submit', e=>{
      e.preventDefault(); clearError();

      // 必須チェック（最小限）
      if (!f.CustomerID.value){ showError('お客様IDは必須です。'); return; }
      if (!f.Plan.value){ showError('プランを選択してください。'); return; }
      if (!f.Start.value){ showError('開始日を入力してください。'); return; }

      // 申込日を当日で自動セット（hidden）
      const t = new Date();
      const y=t.getFullYear(), m=String(t.getMonth()+1).padStart(2,'0'), d=String(t.getDate()).padStart(2,'0');
      f.ApplyDate.value = `${y}-${m}-${d}`;

      // ★ 送信直前ミラー
      syncHiddenBeforeSubmit();

      const fd=new FormData(e.target); const payload={}; for(const [k,v] of fd.entries()) payload[k]=String(v??'');
      
      // Ensure SameAsTraveler checkbox value is always sent
      payload.SameAsTraveler = f.SameAsTraveler.checked ? 'on' : 'off';
      
      // Debug: Log the checkbox state
      console.log('SameAsTraveler checkbox state:', f.SameAsTraveler.checked);
      console.log('SameAsTraveler value being sent:', payload.SameAsTraveler);

      setLoading(true);
      google.script.run
        .withSuccessHandler(res=>{
          setLoading(false);
          if(res&&res.ok){ window.location.search='?view=thanks&pdfUrl='+encodeURIComponent(res.pdfUrl);}
          else { showError('エラー: '+(res&&res.error||'Unknown')); }
        })
        .withFailureHandler(err=>{
          setLoading(false);
          showError('通信エラー: '+(err&&err.message||err));
        })
        .doPost(payload);
    });

    document.getElementById('resetBtn').addEventListener('click', ()=>{
      clearError();
      f.eligStatus.textContent='初期化しました。入力を開始してください。';
      lockPlan(true);
      f.CoverageLength.selectedIndex=0; f.PlanMeta.textContent=''; f.Price.value=''; f.End.value='';
      f.IllnessSearch.value=''; renderIllnessOptions(ILLNESS_ALL);
      document.getElementById('Q1_Detail').style.display='none';
      document.getElementById('Q2_Detail').style.display='none';
      document.getElementById('Q4_Detail').style.display='none';
      document.getElementById('Q8_Detail').style.display='none';
      document.getElementById('Q5_OtherWrap').style.display='none';

      // hidden もクリアしておくと安全
      document.getElementById('PlanHidden').value = '';
      document.getElementById('CoverageStartHidden').value = '';
      document.getElementById('CoverageEndHidden').value = '';
      document.getElementById('PriceHidden').value = '';
    });

    function escapeHtml(s){ return String(s||'').replace(/[&<>"']/g, m=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[m])); }

    // ===== Boot: シートからルール読込 =====
    function reportRuleCounts(){
      const msg=`Loaded: plans=${(RULES.plans||[]).length}, illnesses=${(RULES.illnesses||[]).length}, durationHeaders=${(RULES.durationHeaders||[]).length}`;
      f.ruleCounts.textContent = msg.replace('Loaded: ','');
      if (!(RULES.durationHeaders||[]).length) f.eligStatus.textContent='期間ヘッダが見つかりません。';
    }
    function boot(){
      google.script.run
        .withSuccessHandler(res=>{
          if(!res||!res.ok){ showError('ルールデータの読込に失敗しました。'); return; }
          RULES=res.data||RULES;
          buildDurationOptions();
          buildIllnessList();
          loadSchoolList();
          // イベント
          f.Age.addEventListener('input', updateEligibility);
          f.Illness.addEventListener('change', ()=>{ f.Illness.setAttribute('data-current', f.Illness.value); updateEligibility(); });
          f.IllnessSearch.addEventListener('input', (e)=>filterIllnessOptions(e.target.value));
          f.Dangerous.addEventListener('change', updateEligibility);
          f.CoverageLength.addEventListener('change', updateEligibility);
          f.Plan.addEventListener('change', updatePriceMeta);
          f.Start.addEventListener('change', computeEndDate);

          reportRuleCounts();
          updateEligibility();

          // 初期 ApplyDate セット
          const t = new Date();
          const y=t.getFullYear(), m=String(t.getMonth()+1).padStart(2,'0'), d=String(t.getDate()).padStart(2,'0');
          f.ApplyDate.value = `${y}-${m}-${d}`;
        })
        .withFailureHandler(err=>{ showError('初期化エラー: '+(err&&err.message||err)); })
        .getEligibilityData();
    }
    boot();
  </script>
</body>
</html>
