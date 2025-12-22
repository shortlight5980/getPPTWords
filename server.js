// 30 行代码搞定 PPTX 文本 + SmartArt 文字提取
const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const JSZip   = require('jszip');
const fs      = require('fs');
const path    = require('path');

const app = express();
app.use(cors());
const upload = multer({ dest: 'uploads/' });

// 通用函数：从任意 xml 里把 <a:t>xxx</a:t> 扫出来
// 1. 把文件内容读成字符串（异步）
async function extractTexts(zip, xmlPath) {
  const file = zip.file(xmlPath);
  if (!file) return [];
  const xml = await file.async('text');

  const out = [];

  // 1. 优先按完整段落拼
  const pRe = /<a:p(?:\s[^>]*)?>(.*?)<\/a:p>/gs;
  const tRe = /<a:t[^>]*>([^<]*)<\/a:t>/g;
  let pMatch;
  while ((pMatch = pRe.exec(xml)) !== null) {
    const segment = pMatch[1];
    let tMatch, line = [];
    while ((tMatch = tRe.exec(segment)) !== null) line.push(tMatch[1]);
    if (line.length) out.push(line.join(''));
    tRe.lastIndex = 0;          // 重置，供下一段使用
  }

  // 2. 如果根本没有 <a:p>，再退回到裸 <a:t>
  if (out.length === 0) {
    tRe.lastIndex = 0;
    let tMatch;
    while ((tMatch = tRe.exec(xml)) !== null) out.push(tMatch[1]);
  }

  return out;
}

// 新增：把 rId -> 目标文件 建一张表
async function buildRelMap(zip, slideIdx) {
  const relPath = `ppt/slides/_rels/slide${slideIdx}.xml.rels`;
  const file = zip.file(relPath);
  if (!file) return {};
  const xml = await file.async('text');
  const map = {};
  // 例：<Relationship Id="rId3" Type=".../chart" Target="../charts/chart2.xml"/>
  const re = /Id="([^"]+)"[^T]*Target="([^"]+)"/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    map[m[1]] = m[2];          // rId -> charts/chartN.xml
  }
  return map;
}

// 新增：根据 slide 拿到它里面所有 chart 的文字
async function extractChartTexts(zip, slideIdx) {
  const relMap = await buildRelMap(zip, slideIdx);
  const texts = [];
  for (const [rId, target] of Object.entries(relMap)) {
    if (target.includes('/charts/')) {
      const chartTexts = await extractTexts(zip, target); // 复用老函数
      texts.push(...chartTexts);
    }
  }
  return texts;
}

app.post('/ppt', upload.single('ppt'), async (req, res) => {
  try {
    const buf = fs.readFileSync(req.file.path);
    const zip = await JSZip.loadAsync(buf);

    const slideList = Object.keys(zip.files)
      .filter(f => /^ppt\/slides\/slide(\d+)\.xml$/.test(f))
      .sort((a, b) => a.match(/(\d+)/)[1] - b.match(/(\d+)/)[1]);

    const result = [];
    for (const slidePath of slideList) {
      const idx = +slidePath.match(/slide(\d+)\.xml/)[1];
      const normalTexts = await extractTexts(zip, slidePath);
      const smartTexts  = await extractTexts(zip, `ppt/diagrams/data${idx}.xml`);
      const chartTexts  = await extractChartTexts(zip, idx);   // 新增
      result.push({
        slide: idx,
        texts: [...normalTexts, ...smartTexts, ...chartTexts] // 合并三类
      });
    }
    res.json(result);
  } catch (e) {
    res.status(500).json({ error: e.message });
  } finally {
    if (req.file) fs.unlinkSync(req.file.path);
  }
});

app.listen(3000, () => console.log('>>>  http://localhost:3000  <<< 已启动'));