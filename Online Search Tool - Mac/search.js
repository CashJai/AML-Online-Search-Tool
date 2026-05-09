const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const XLSX = require('xlsx-js-style');
const { PDFDocument, rgb } = require('pdf-lib');
const cc = require('chinese-conv');

// --- CONSTANTS & DICTIONARIES ---
const ADVERSE_KEYWORDS = `"money laundering" OR "fraud" OR "corruption" OR "bribery" OR "sanctions" OR "OFAC" OR "regulatory fine" OR "criminal charges" OR "indictment" OR "conviction" OR "financial crime" OR "illegal activity" OR "misconduct" OR "illicit" OR "unethical" OR "controversy" OR "arrested" OR "charged" OR "investigation" OR "lawsuit" OR "penalty" OR "politically exposed person" OR "PEP" OR "influence peddling" OR "abuse of power" OR "tax evasion" OR "Terrorist" OR "Trafficking" OR "Narcotics"`;

const FUZZY_DICT = {
  chan: ['chen'], chen: ['chan'],
  zhang: ['chang', 'cheung'], chang: ['zhang', 'cheung'], cheung: ['zhang', 'chang'],
  man: ['men'], men: ['man'],
  mei: ['may'], may: ['mei'],
  zhou: ['chou', 'chow'], chou: ['zhou', 'chow'], chow: ['zhou', 'chou'],
  lee: ['li'], li: ['lee'],
  wong: ['wang'], wang: ['wong'],
  wu: ['ng'], ng: ['wu'],
  lin: ['lam'], lam: ['lin'],
  zheng: ['cheng'], cheng: ['zheng'],
  xu: ['hsu', 'hui'], hsu: ['xu', 'hui'], hui: ['xu', 'hsu'],
  cai: ['tsai', 'choy', 'choi'], tsai: ['cai', 'choy', 'choi'], choy: ['cai', 'tsai', 'choi'], choi: ['cai', 'tsai', 'choy'],
  guo: ['kuo', 'kwok'], kuo: ['guo', 'kwok'], kwok: ['guo', 'kuo'],
  huang: ['hwang'], hwang: ['huang'],
  zhao: ['chao', 'chiu'], chao: ['zhao', 'chiu'], chiu: ['zhao', 'chao']
};

const NAME_STOPWORDS = new Set([
  "the", "a", "an", "in", "on", "at", "to", "and", "or", "but", "is", "are", 
  "was", "were", "has", "have", "had", "mr", "mrs", "ms", "dr", "president", 
  "ceo", "director", "manager", "officer", "judge", "justice", "mayor", 
  "governor", "senator", "representative", "defendant", "plaintiff", "ltd", 
  "inc", "llc", "corp", "company", "bank", "group", "of", "for", "by", "with",
  "holdings", "international", "co", "partners", "limited", "pte", "corporation"
]);

const HIGH_RISK_OVERRIDE = [
  /arrest/i, /charg/i, /indict/i, /convict/i, /guilty/i, /prison/i, /jail/i, /sentenc/i, /fine/i, /penal/i, /sued/i,
  /逮捕/, /拘捕/, /起诉/, /定罪/, /有罪/, /坐牢/, /监禁/, /判刑/, /罚款/, /惩罚/, /诉讼/
];

const FALSE_POSITIVES = [
  {
    regex1: /\b(acams|acfcs|fincen|fatf|sfc|sec|icac|doj)\b/gi,
    regex2: /\b(event|webinar|roundtable|scholarship|certification|circular|guidelines|unveils|framework)\b/gi,
    dist: 80
  },
  {
    regex1: /\b(anti-money laundering|aml|kyc|compliance)\b/gi,
    regex2: /\b(policy|framework|controls|procedures|training)\b/gi,
    dist: 60
  },
  {
    regex1: /\b(ordinance|law|act|statute)\b/gi,
    regex2: /\b(money laundering|fraud|bribery)\b/gi,
    dist: 40
  },
  {
    regex1: /\b(investigation)\b/gi,
    regex2: /\b(disease|virus|cancer|cognitive|machine learning|physics)\b/gi,
    dist: 60
  },
  {
    // Prevention & Detection
    regex1: /\b(prevent|detect|combat|fight|tackle|stop|raise awareness|protect against|mitigate|solution|software|tool)\b/gi,
    regex2: /\b(money laundering|fraud|scam|corruption|bribery|crime|illicit)\b/gi,
    dist: 60
  },
  {
    // Chinese Prevention
    regex1: /(防范|侦测|预防|打击|遏制|意识|解决方案|软件|系统|工具)/g,
    regex2: /(洗钱|洗黑钱|诈骗|欺诈|贪污|腐败|金融犯罪|犯罪)/g,
    dist: 40
  },
  {
    // Education & Training
    regex1: /\b(exam|certification|training|course|seminar|webinar|conference|summit|workshop|guide|best practices)\b/gi,
    regex2: /\b(money laundering|aml|fraud|compliance|sanctions)\b/gi,
    dist: 60
  },
  {
    // Chinese Education
    regex1: /(考试|认证|培训|课程|研讨会|峰会|指南|最佳实践|论坛)/g,
    regex2: /(洗钱|反洗钱|诈骗|合规|制裁)/g,
    dist: 40
  },
  {
    // Regulatory framework
    regex1: /\b(policy|framework|controls|procedures|ordinance|law|act|statute|regulation)\b/gi,
    regex2: /\b(money laundering|fraud|bribery|sanctions)\b/gi,
    dist: 60
  },
  {
    // Chinese Regulatory
    regex1: /(政策|框架|控制|程序|条例|法律|法案|法规|合规)/g,
    regex2: /(洗钱|诈骗|贪污|制裁)/g,
    dist: 40
  }
];

const CRIME_PATTERNS = [
  { regex: /money\s*launder/i, label: "Money Laundering" },
  { regex: /fraud/i, label: "Fraud" },
  { regex: /corrupt/i, label: "Corruption" },
  { regex: /brib/i, label: "Bribery" },
  { regex: /sanction/i, label: "Sanctions" },
  { regex: /\bofac\b/i, label: "OFAC" },
  { regex: /regulatory\s*fine/i, label: "Regulatory Fine" },
  { regex: /criminal/i, label: "Criminal" },
  { regex: /indict/i, label: "Indictment" },
  { regex: /convict/i, label: "Conviction" },
  { regex: /financial\s*crime/i, label: "Financial Crime" },
  { regex: /illegal/i, label: "Illegal Activity" },
  { regex: /misconduct/i, label: "Misconduct" },
  { regex: /illicit/i, label: "Illicit" },
  { regex: /unethical/i, label: "Unethical" },
  { regex: /controversy/i, label: "Controversy" },
  { regex: /arrest/i, label: "Arrest" },
  { regex: /charg/i, label: "Charged" },
  { regex: /investigat/i, label: "Investigation" },
  { regex: /lawsuit/i, label: "Lawsuit" },
  { regex: /sued/i, label: "Sued" },
  { regex: /penal/i, label: "Penalty" },
  { regex: /politically\s*exposed/i, label: "PEP" },
  { regex: /\bpep\b/i, label: "PEP" },
  { regex: /influence\s*peddling/i, label: "Influence Peddling" },
  { regex: /abuse\s*of\s*power/i, label: "Abuse of Power" },
  { regex: /tax\s*evasion/i, label: "Tax Evasion" },
  { regex: /tax\s*fraud/i, label: "Tax Fraud" },
  { regex: /terror/i, label: "Terrorism" },
  { regex: /traffick/i, label: "Trafficking" },
  { regex: /narcotic/i, label: "Narcotics" },
  { regex: /guilty/i, label: "Guilty" },
  { regex: /prison/i, label: "Prison" },
  { regex: /jail/i, label: "Jail" },
  { regex: /sentenc/i, label: "Sentenced" },
  { regex: /embezzle/i, label: "Embezzlement" },
  { regex: /extortion/i, label: "Extortion" },
  { regex: /scam/i, label: "Scam" },
  { regex: /ponzi/i, label: "Ponzi" },
  { regex: /insider\s*trading/i, label: "Insider Trading" },
  { regex: /market\s*manipulation/i, label: "Market Manipulation" },
  { regex: /forgery/i, label: "Forgery" },
  { regex: /smuggling/i, label: "Smuggling" },
  // Chinese Terms
  { regex: /洗钱|洗黑钱/, label: "Money Laundering (洗錢)" },
  { regex: /欺诈|诈骗/, label: "Fraud (詐騙)" },
  { regex: /贪污|腐败|贿赂/, label: "Corruption/Bribery (貪污/賄賂)" },
  { regex: /制裁/, label: "Sanctions (制裁)" },
  { regex: /罚款|惩罚/, label: "Penalty/Fine (罰款)" },
  { regex: /刑事|起诉|定罪|控告|指控|有罪|判刑/, label: "Criminal/Charged/Convicted (刑事/起訴)" },
  { regex: /金融犯罪/, label: "Financial Crime (金融犯罪)" },
  { regex: /非法|违规/, label: "Illegal (非法)" },
  { regex: /不当行为/, label: "Misconduct (不當行為)" },
  { regex: /争议/, label: "Controversy (爭議)" },
  { regex: /逮捕|拘捕|坐牢|监禁/, label: "Arrested/Jail (逮捕)" },
  { regex: /调查|廉政公署|证监会|金管局/, label: "Investigation/Regulator (調查)" },
  { regex: /诉讼/, label: "Lawsuit (訴訟)" },
  { regex: /政治人物/, label: "PEP (政治人物)" },
  { regex: /逃税/, label: "Tax Evasion (逃稅)" },
  { regex: /恐怖主义/, label: "Terrorism (恐怖主義)" },
  { regex: /贩卖|毒品|走私/, label: "Trafficking/Narcotics (販賣/毒品)" },
  { regex: /挪用公款/, label: "Embezzlement (挪用公款)" },
  { regex: /勒索/, label: "Extortion (勒索)" },
  { regex: /骗局/, label: "Scam (騙局)" },
  { regex: /内幕交易/, label: "Insider Trading (內幕交易)" },
  { regex: /操纵市场/, label: "Market Manipulation (操縱市場)" },
  { regex: /伪造/, label: "Forgery (偽造)" }
];

// --- HELPER FUNCTIONS ---
function stripLegalSuffixes(name) {
  const suffixes = /(?:\s+|,)(ltd|llc|corporation|inc|limited|company|pte|co|股份有限公司|有限公司|公司)\.?\b/gi;
  return name.replace(suffixes, '').trim();
}

function normalize(text) {
  if (!text) return '';
  return cc.tify(text).toLowerCase();
}

function isFalsePositive(snippet) {
  const norm = snippet.toLowerCase();

  // If there's an explicit high-risk action, do not treat as false positive
  for (const override of HIGH_RISK_OVERRIDE) {
    if (override.test(norm)) return false;
  }

  for (const fp of FALSE_POSITIVES) {
    const matches1 = [...norm.matchAll(fp.regex1)];
    const matches2 = [...norm.matchAll(fp.regex2)];
    
    for (const m1 of matches1) {
      for (const m2 of matches2) {
        if (Math.abs(m1.index - m2.index) <= fp.dist) {
          return true; // Is a false positive
        }
      }
    }
  }
  return false;
}

function generateAliases(subjectTokens) {
  let results = [[]];
  for (const token of subjectTokens) {
    const t = token.toLowerCase();
    const options = [t];
    if (FUZZY_DICT[t]) {
      options.push(...FUZZY_DICT[t]);
    }
    const newResults = [];
    for (const res of results) {
      for (const opt of options) {
        newResults.push([...res, opt]);
      }
    }
    results = newResults;
  }
  const aliases = results.map(arr => arr.join(" "));
  
  if (subjectTokens.length > 2) {
    const firstTokenOptions = [subjectTokens[0].toLowerCase()];
    if (FUZZY_DICT[firstTokenOptions[0]]) firstTokenOptions.push(...FUZZY_DICT[firstTokenOptions[0]]);
    
    const lastTokenOptions = [subjectTokens[subjectTokens.length - 1].toLowerCase()];
    if (FUZZY_DICT[lastTokenOptions[0]]) lastTokenOptions.push(...FUZZY_DICT[lastTokenOptions[0]]);
    
    for (const f of firstTokenOptions) {
      for (const l of lastTokenOptions) {
        aliases.push(`${f} ${l}`);
      }
    }
  }
  return aliases;
}

function isExactAliasMatch(validAliases, originalText) {
  for (const alias of validAliases) {
    const escapedAlias = alias.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp('\\b' + escapedAlias + '\\b', 'gi');
    
    let match;
    while ((match = regex.exec(originalText)) !== null) {
      const matchEnd = regex.lastIndex;
      let isMismatch = false;
      
      const afterText = originalText.substring(matchEnd);
      const afterMatch = afterText.match(/^\s+([A-Za-z]+)/);
      if (afterMatch) {
        const wordAfter = afterMatch[1];
        if (/^[A-Z][a-z]*$/.test(wordAfter)) {
          if (!NAME_STOPWORDS.has(wordAfter.toLowerCase())) {
            isMismatch = true;
          }
        }
      }
      
      if (!isMismatch) {
        return true; 
      }
    }
  }
  return false;
}

function evaluateMatch(subject, coreName, title, snippet, url) {
  const normUrl = url.toLowerCase();
  const originalText = title + " " + snippet;

  const subjectTokens = subject.split(/\s+/);
  const subjectAliases = generateAliases(subjectTokens);
  
  // 1. Exact / Alias Match (Strict Boundary Check)
  if (isExactAliasMatch(subjectAliases, originalText)) {
    return "Exact Match";
  }

  // 2. Likely Match (Core Name Strict Boundary Check)
  if (coreName) {
    const coreTokens = coreName.split(/\s+/);
    const coreAliases = generateAliases(coreTokens);
    if (isExactAliasMatch(coreAliases, originalText)) {
      return "Likely Match";
    }
  }

  // 3. URL Match
  const coreTokens = coreName.split(/\s+/).map(t => t.toLowerCase());
  let allTokensInUrl = true;
  for (const t of coreTokens) {
    if (!normUrl.includes(t)) {
      allTokensInUrl = false;
      break;
    }
  }
  if (coreTokens.length > 0 && allTokensInUrl) {
    return "URL Match";
  }

  return "Not a Match";
}

function evaluateFCC(title, snippet) {
  const combinedText = normalize(title + " " + snippet);
  
  if (isFalsePositive(combinedText)) {
    return { isFCC: false, evidence: "" };
  }

  const foundKeywords = [];
  for (const pattern of CRIME_PATTERNS) {
    if (pattern.regex.test(combinedText)) {
      foundKeywords.push(pattern.label);
    }
  }

  if (foundKeywords.length > 0) {
    const uniqueKeywords = [...new Set(foundKeywords)];
    return {
      isFCC: true,
      evidence: `FCC Concern: subject name found in title/snippet and FCC-related keywords detected. Keywords: ${uniqueKeywords.join(", ")}`
    };
  }

  return { isFCC: false, evidence: "" };
}

// --- PDF GENERATION ---
async function createPDF(images, outputPath, timestamp, query) {
  const pdfDoc = await PDFDocument.create();
  
  for (const imgBuffer of images) {
    const page = pdfDoc.addPage();
    const { width, height } = page.getSize();
    
    try {
      const image = await pdfDoc.embedPng(imgBuffer);
      const imgDims = image.scale(1);
      
      const targetWidth = width - 40;
      const scale = targetWidth / imgDims.width;
      const drawWidth = imgDims.width * scale;
      const drawHeight = imgDims.height * scale;

      // Adjust page size to image + header margin
      page.setSize(drawWidth + 40, drawHeight + 60);
      
      // Draw banner
      page.drawText(`Search Time: ${timestamp} | Query: ${query}`, {
        x: 20,
        y: drawHeight + 40,
        size: 10,
        color: rgb(0, 0, 0)
      });

      page.drawImage(image, {
        x: 20,
        y: 20,
        width: drawWidth,
        height: drawHeight
      });
    } catch (e) {
      console.error("Failed to embed image in PDF:", e);
    }
  }

  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, pdfBytes);
}

// --- EXCEL GENERATION ---
function generateExcel(allSubjectData, outputPath) {
  const wb = XLSX.utils.book_new();

  // Helper to create a results sheet with summary header
  function createResultsSheet(sheetResults, sheetName, meta, searchTime) {
    const tabRisk = sheetResults.some(r => r.riskLevel === "High") ? "High" : "Low";
    
    const wsData = [
      ["Search Term", meta.subject],
      ["Engine", "Bing"],
      ["Region of Search", "Hong Kong"],
      ["Search Time", searchTime],
      ["Overall Risk", tabRisk],
      [], // empty row
      ["Page", "Rank", "Result Title", "Result Snippet", "Result URL", "Subject Matching", "FCC Concern", "Risk Level"]
    ];

    sheetResults.forEach(r => {
      wsData.push([
        r.page, r.rank, r.title, r.snippet, r.url, r.matchType, r.fccEvidence || "None", r.riskLevel
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Style summary labels (rows 0-4, column 0) bold
    for (let r = 0; r <= 4; r++) {
      const cellRef = XLSX.utils.encode_cell({r, c: 0});
      if (ws[cellRef]) {
        ws[cellRef].s = { font: { bold: true } };
      }
    }

    // Highlight risk value if High
    const riskCell = XLSX.utils.encode_cell({r: 4, c: 1});
    if (ws[riskCell] && tabRisk === "High") {
      ws[riskCell].s = { fill: { fgColor: { rgb: "FFC7CE" } }, font: { color: { rgb: "9C0006" }, bold: true } };
    }

    // Style table header row (row 6)
    const headerRow = 6;
    for (let c = 0; c < 8; c++) {
      const cellRef = XLSX.utils.encode_cell({r: headerRow, c});
      if (!ws[cellRef]) ws[cellRef] = { v: wsData[headerRow][c], t: 's' };
      ws[cellRef].s = {
        fill: { fgColor: { rgb: "0000FF" } },
        font: { color: { rgb: "FFFFFF" }, bold: true }
      };
    }

    // Style data rows
    sheetResults.forEach((r, idx) => {
      const rIdx = idx + 7; // data starts at row 7
      if (r.riskLevel === "High") {
        for (let c = 0; c < 8; c++) {
          const cellRef = XLSX.utils.encode_cell({r: rIdx, c});
          if (ws[cellRef]) {
            ws[cellRef].s = ws[cellRef].s || {};
            ws[cellRef].s.fill = { fgColor: { rgb: "FFC7CE" } };
            if (c === 7) {
              ws[cellRef].s.font = { color: { rgb: "9C0006" }, bold: true };
            }
          }
        }
      }

      // URL styling
      const urlRef = XLSX.utils.encode_cell({r: rIdx, c: 4});
      if (ws[urlRef]) {
        ws[urlRef].s = ws[urlRef].s || {};
        ws[urlRef].s.font = { color: { rgb: "0000FF" }, underline: true };
      }
    });

    ws['!cols'] = [
      {wch: 12}, {wch: 6}, {wch: 40}, {wch: 60}, {wch: 40}, {wch: 15}, {wch: 40}, {wch: 10}
    ];

    // Excel tab names max 31 chars
    const safeName = sheetName.substring(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  }

  // Create two tabs per subject
  for (const data of allSubjectData) {
    const { subject, results, meta } = data;
    const label = subject.length > 20 ? subject.substring(0, 20) : subject;

    const onlineResults = results.filter(r => r.type === 'Online');
    const adverseResults = results.filter(r => r.type === 'Adverse');

    createResultsSheet(onlineResults, `${label} - Online`, meta, meta.genTime);
    createResultsSheet(adverseResults, `${label} - Adverse`, meta, meta.advTime);
  }

  XLSX.writeFile(wb, outputPath);
}

// --- MAIN SEARCH LOGIC ---
async function runSearchEngine(subjects) {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    viewport: { width: 1920, height: 1080 },
    locale: 'en-HK'
  });

  // Create one shared output folder for this entire run
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 16);
  const resultsBaseDir = path.join(process.cwd(), 'Search Results');
  if (!fs.existsSync(resultsBaseDir)) {
    fs.mkdirSync(resultsBaseDir);
  }
  const outputDir = path.join(resultsBaseDir, `Search_${timestamp}`);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }

  // Collect all subjects' data for the single Excel file
  const allSubjectData = [];
  
  for (const subject of subjects) {
    const coreName = stripLegalSuffixes(subject);
    const safeSubjectName = subject.replace(/[^a-zA-Z0-9]/g, '_');

    const allResults = [];
    const meta = { subject, overallRisk: "Low", genTime: "", advTime: "" };

    const searches = [
      { type: 'Online', query: subject },
      { type: 'Adverse', query: `"${subject}" AND (${ADVERSE_KEYWORDS})` }
    ];

    for (const search of searches) {
      const page = await context.newPage();
      const images = [];
      let currentTimestamp = new Date().toISOString();
      if (search.type === 'Online') meta.genTime = currentTimestamp;
      else meta.advTime = currentTimestamp;

      console.log(`\nStarting ${search.type} Search for: ${subject}`);
      console.log(`Query: ${search.query}`);
      
      // Navigate to Bing
      await context.clearCookies();
      const bingUrl = `https://www.bing.com/search?q=${encodeURIComponent(search.query)}`;
      await page.goto(bingUrl, { waitUntil: 'domcontentloaded' });
      
      console.log("Waiting for Bing results...");
      await page.locator('li.b_algo').first().waitFor({ state: 'visible', timeout: 30000 }).catch(() => {});

      // Scrape 3 pages
      let rankCounter = 1;
      for (let pNum = 1; pNum <= 3; pNum++) {
        console.log(`  Scraping page ${pNum}...`);
        await page.waitForTimeout(2000); 
        
        // Full page screenshot
        const imgBuffer = await page.screenshot({ fullPage: true });
        images.push(imgBuffer);

        // Extract results
        const resultsLocators = page.locator('li.b_algo');
        const count = await resultsLocators.count();
        
        if (count === 0) {
           console.log("  No results found on this page.");
           break;
        }

        for (let i = 0; i < count; i++) {
          const res = resultsLocators.nth(i);
          const titleElem = res.locator('h2 a').first();
          const title = await titleElem.innerText().catch(() => "");
          let url = await titleElem.getAttribute('href').catch(() => "");

          if (url && url.includes('/ck/a')) {
            const uAttr = await titleElem.getAttribute('u').catch(() => null);
            if (uAttr && uAttr.startsWith('a1')) {
              try {
                const decodedUrl = Buffer.from(uAttr.substring(2), 'base64').toString('utf8');
                if (decodedUrl.startsWith('http')) url = decodedUrl;
              } catch (e) {}
            }
            
            if (url.includes('/ck/a')) {
              const citeElem = res.locator('cite').first();
              if (await citeElem.count() > 0) {
                const citeText = await citeElem.innerText().catch(() => "");
                const cleanCite = citeText.split(' ')[0].split('›')[0].trim();
                if (cleanCite.startsWith('http')) {
                  url = cleanCite;
                } else if (cleanCite) {
                  url = "https://" + cleanCite;
                }
              }
            }
          }

          let snippet = "";
          const snippetElem = res.locator('.b_caption p, .b_algoSlug, .b_snippet, .b_paractl, p').first();
          if (await snippetElem.count() > 0) {
            snippet = await snippetElem.innerText().catch(() => "");
          }

          if (title && url) {
            const matchType = evaluateMatch(subject, coreName, title, snippet, url);
            let fccEvidence = "";
            let riskLevel = "Low";
            
            if (search.type === 'Adverse') {
              const fccEval = evaluateFCC(title, snippet);
              fccEvidence = fccEval.evidence;
              if (matchType !== "Not a Match" && fccEval.isFCC) {
                riskLevel = "High";
                meta.overallRisk = "High";
              }
            }

            allResults.push({
              type: search.type,
              page: pNum,
              rank: rankCounter++,
              title: title.replace(/\n/g, ' '),
              snippet: snippet.replace(/\n/g, ' '),
              url,
              matchType,
              fccEvidence,
              riskLevel
            });
          }
        }

        // Next page
        if (pNum < 3) {
          const nextBtn = page.locator('a.sb_pagN, a[title="Next page"], a.b_widePag, .b_pag a').first();
          if (await nextBtn.isVisible()) {
            await nextBtn.click();
            await page.waitForLoadState('domcontentloaded');
            await page.locator('li.b_algo').first().waitFor({ state: 'visible', timeout: 30000 }).catch(() => {});
          } else {
            console.log("  No more pages available.");
            break;
          }
        }
      }
      
      const pdfPath = path.join(outputDir, `${safeSubjectName}_${search.type}_Search.pdf`);
      await createPDF(images, pdfPath, currentTimestamp, search.query);
      console.log(`Saved PDF: ${pdfPath}`);
      
      await page.close();
    }

    // Collect this subject's data for the combined Excel
    allSubjectData.push({ subject, results: allResults, meta });
  }

  // Generate one single Excel file with all subjects
  const excelPath = path.join(outputDir, `Search_Results.xlsx`);
  generateExcel(allSubjectData, excelPath);
  console.log(`\nSaved Excel: ${excelPath}`);

  await browser.close();
}

// Execution
const args = process.argv.slice(2);
if (args.length === 0) {
  console.log("No subjects provided.");
  process.exit(1);
}

runSearchEngine(args).then(() => {
  console.log("\nAll searches completed successfully.");
}).catch(err => {
  console.error("\nError during search:", err);
});

