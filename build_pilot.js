/**
 * build_pilot.js
 * Builds pilot Word document (pages 1-5) for Chapter 4 of Chinese calculus textbook.
 * Re-runnable: node build_pilot.js
 * 
 * Strategy:
 *  - Pages 1-2 (cover/chapter title): embed as full-page images
 *  - Pages 3-5: transcribe Chinese text, embed math as cropped images
 */

const {
  Document, Packer, Paragraph, TextRun, ImageRun,
  HeadingLevel, AlignmentType, LevelFormat,
  BorderStyle, WidthType, PageOrientation,
  ShadingType, VerticalAlign
} = require('docx');

const fs = require('fs');
const path = require('path');

const CROPS = '/home/claude/math_crops';
const OUT   = '/mnt/user-data/outputs/chapter4_pilot.docx';

// Helper: load image buffer + dimensions
function img(name) {
  const buf = fs.readFileSync(path.join(CROPS, name + '.png'));
  // Get dimensions from PNG header
  const w = buf.readUInt32BE(16);
  const h = buf.readUInt32BE(20);
  return { buf, w, h };
}

// Helper: create a centred image paragraph, scaled to fit content width
// contentWidthEMU = 6.5 inches * 914400 = 5943600
function imgPara(name, maxWidthIn = 5.5, align = AlignmentType.CENTER) {
  const { buf, w, h } = img(name);
  const maxW = maxWidthIn * 914400;
  const scale = Math.min(1, maxW / (w * 914400 / 96)); // assume 96dpi source
  const emuW = Math.round(w * 914400 / 96 * scale);
  const emuH = Math.round(h * 914400 / 96 * scale);
  return new Paragraph({
    alignment: align,
    spacing: { before: 80, after: 80 },
    children: [
      new ImageRun({ data: buf, transformation: { width: emuW, height: emuH }, type: 'png' })
    ]
  });
}

// Helper: inline image (for math within a sentence)
function inlineImg(name, heightIn = 0.22) {
  const { buf, w, h } = img(name);
  const emuH = Math.round(heightIn * 914400);
  const emuW = Math.round(emuH * w / h);
  return new ImageRun({ data: buf, transformation: { width: emuW, height: emuH }, type: 'png' });
}

// Paragraph styles
const bodyFont = 'SimSun';  // standard Chinese serif font (fallback: Times New Roman)
const bodySize = 22;        // 11pt in half-points
const spacing  = { before: 0, after: 120, line: 276 };

function body(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing,
    indent: opts.indent !== false ? { firstLine: 440 } : {},  // ~2 Chinese chars indent
    children: [new TextRun({ text, font: bodyFont, size: bodySize, ...opts.run })]
  });
}

function bodyRuns(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing,
    indent: opts.indent !== false ? { firstLine: 440 } : {},
    children: runs
  });
}

function tr(text, opts = {}) {
  return new TextRun({ text, font: bodyFont, size: bodySize, ...opts });
}

function heading2(text) {
  // Section heading style: bold, larger, e.g. "4.1 再论极限——罗必塔法则"
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 180 },
    indent: {},
    children: [new TextRun({ text, font: 'SimHei', size: 28, bold: true })]
  });
}

function subheading(text) {
  // Bold sub-heading within section, e.g. "罗必塔法则(L`Hôpital`s Rule)"
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    indent: { firstLine: 440 },
    children: [new TextRun({ text, font: 'SimHei', size: 24, bold: true })]
  });
}

function example(label, rest) {
  // Bold "例4-1)" label followed by regular text
  return new Paragraph({
    spacing,
    indent: { firstLine: 440 },
    children: [
      new TextRun({ text: label, font: bodyFont, size: bodySize, bold: true }),
      new TextRun({ text: '  ' + rest, font: bodyFont, size: bodySize })
    ]
  });
}

function solution(text) {
  return new Paragraph({
    spacing,
    indent: { firstLine: 440 },
    children: [new TextRun({ text, font: bodyFont, size: bodySize })]
  });
}

function blank(pts = 80) {
  return new Paragraph({ spacing: { before: 0, after: pts }, children: [] });
}

// ──────────────────────────────────────────────
// Build content
// ──────────────────────────────────────────────
const children = [];

// ── Page 1: Table of Contents cover ──
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: '[Page 1 — Table of Contents]', font: 'Arial', size: 20, color: '888888', italics: true })]
  }),
  imgPara('cover_toc', 5.5),
  new Paragraph({ pageBreakBefore: true, children: [] })
);

// ── Page 2: Chapter title page ──
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: '[Page 2 — Chapter Title Page]', font: 'Arial', size: 20, color: '888888', italics: true })]
  }),
  imgPara('cover_chapter', 5.5),
  new Paragraph({ pageBreakBefore: true, children: [] })
);

// ── Page 3: Section 4.1 opening + L'Hopital's Rule ──
// Running header note
children.push(
  new Paragraph({
    spacing: { before: 0, after: 60 },
    indent: {},
    children: [
      new TextRun({ text: '84', font: 'Arial', size: 18, color: '666666' }),
      new TextRun({ text: '                                                                   玩转微积分', font: bodyFont, size: 18, color: '666666' })
    ]
  })
);

// Opening paragraphs
children.push(
  body('前两章，我们弄清楚了微积分中最重要的极限概念，又升华了中学就学过的斜率概念，引入了导数。'),
  body('我们还耗费了大量笔墨，努力搞清楚了各种函数的小趋势导数的求法，这导数总得有点特别的用途，才对得起我们的心血吧？'),
  body('没错，导数真的很有用！'),
  blank()
);

// Section 4.1 heading
children.push(heading2('4.1  再论极限——罗必塔法则（L\'Hôpital\'s Rule）'));

children.push(
  body('如前所述，极限问题是微积分的核心问题，我们在处理极限问题的时候，经常碰到两种棘手的极限问题：'),
  blank(40)
);

// Displayed formula: ∞/∞ 型和 0/0 型
children.push(imgPara('p3_indeterminate', 3.5));
children.push(blank(40));

children.push(
  body('我们看到两个横躺着的"8"，感觉无从下手。'),
  body('正好比，我们看武侠小说，两位大侠比武，他们都身怀绝技，谁能笑到最后，真是不好判断。我们再开动脑筋想一想，两位大侠长期比下去，哪个更占上风？又比如，甲有10亿元并且每年增长7%；乙虽只有5亿元，但是乙的财富增长率是10%。细想一下不难得出结论：长期下来，乙一定是要战胜甲的。我们之所以得出乙胜的结论，是因为他的10%增长率实在是比7%的增长率厉害得多。什么是增长率，不就是变化率吗？什么是变化率呢？导数呗！'),
  body('由此我们想到，如果两个∞相比没有结论，但它们的导数之比有高下的话，那它们本身之比也就有了结果。这个朴素的想法，在300多年前就被法国数学家罗必塔（L\'Hôpital）先生证明了。'),
  blank()
);

// Sub-heading: 罗必塔法则
children.push(subheading('罗必塔法则（L\'Hôpital\'s Rule）'));

children.push(
  body('假设函数 f(x) 和 g(x)，在含有点 a 的某个区域内可导（不要求一定在 a 可导），而且 g(x) ≠ 0。如果：'),
  blank(40)
);

// The theorem formula block
children.push(imgPara('p3_lhopital_formula', 5.0));
children.push(blank(40));

children.push(
  body('当然，前提是右边的极限存在。'),
  body('罗必塔法则，一石二鸟，同时解决了两种难解决的问题。关于这个法则，我们要提醒几点。'),
  blank()
);

// ── Page 4 ──
children.push(
  new Paragraph({
    spacing: { before: 0, after: 60 },
    indent: {},
    children: [
      new TextRun({ text: '第四章│导数有啥用', font: bodyFont, size: 18, color: '666666' }),
      new TextRun({ text: '                                                                          85', font: 'Arial', size: 18, color: '666666' })
    ]
  })
);

children.push(
  body('首先，当 x →±∞ 时，该法则同样适用。'),
  body('其次，如果两个函数的导数之比的极限还是 ∞/∞ 型和 0/0 型，可以再次使用该法则到二阶导数甚至 n 阶导数上。'),
  body('最后，该法则的证明，需要更高深的数学知识。'),
  body('眼下，我们可以把它当成公理，认为它不证自明，会用法则是硬道理。'),
  blank()
);

// Example 4-1
children.push(example('例 4-1）', '求 lim (ln x)/(x - 1)。\u2003（x → 1）'));
children.push(solution('解：根据罗必塔法则，我们有：'));
children.push(blank(40));
children.push(imgPara('p4_ex1_solution', 5.0));
children.push(blank());

// Example 4-2
children.push(example('例 4-2）', '求 lim 2ˣ / 8x。\u2003（x → ∞）'));
children.push(solution('解：用罗必塔法则，得到：'));
children.push(blank(40));
children.push(imgPara('p4_ex2_solution', 5.0));
children.push(blank());

children.push(body('好玩吧？再看两个极限。'));
children.push(blank());

// Example 4-3
children.push(example('例 4-3）', '求 lim 3ˣ / x³。\u2003（x → ∞）'));
children.push(solution('解：求这个极限，得对分子分母分别求3次导数，才能看清楚结果：'));
children.push(blank(40));
children.push(imgPara('p4_ex3_solution', 5.0));
children.push(blank(40));

children.push(
  body('由本题的解题过程可知，即使分母不是 x³，而是 x¹⁰⁰ 或者更高次方，也不会改变极限的结果。同样，只要分子上 aˣ 中的 a > 1，结果不会变，都是 ∞。'),
  body('这也就印证了我们前面说过的：aˣ 型的 ∞ 最厉害，跟 xⁿ 型的 ∞ 没有可比性。我们再看一下 xⁿ 型的 ∞ 遇上 log_b x 型的 ∞，会发生什么。'),
  blank()
);

// Example 4-4
children.push(example('例 4-4）', ''));
children.push(blank(40));
children.push(imgPara('p4_ex4_full', 5.0));
children.push(blank());

// ── Page 5 ──
children.push(
  new Paragraph({
    spacing: { before: 120, after: 60 },
    indent: {},
    children: [
      new TextRun({ text: '86', font: 'Arial', size: 18, color: '666666' }),
      new TextRun({ text: '                                                                   玩转微积分', font: bodyFont, size: 18, color: '666666' })
    ]
  })
);

children.push(
  body('这一次，xⁿ 型的 ∞ 与 log_b x 型的 ∞ 比较结果同样是 ∞！'),
  body('再重申一次，如果碰到 ∞/∞ 型的极限问题，分子分母只留下级别最高的 ∞ "值班"即可，结果不会错。'),
  body('有了罗必塔法则，再来看第二章要求你记住的两个著名极限，将会非常直观：'),
  blank(40)
);

// Two famous limits
children.push(imgPara('p5_famous_limits', 3.5));
children.push(blank(40));

children.push(
  body('我们登山时会发现，登得越高，视野越好。玩数学也一样，玩得越多，玩过的东西看得越清楚。'),
  body('让我们看看导数还有什么用。'),
  blank()
);

// Section 4.2 heading
children.push(heading2('4.2  极值问题（Extremum）'));

children.push(
  body('人们总爱关注离自己稍微远点的、极端点的事物。'),
  body('同样，函数的最大值和最小值问题，自然引起了人们的极大兴趣。因为在生活中，极值太有用了。'),
  body('要解决极值问题，导数可以帮上大忙。先厘清几个关于极值问题的概念。'),
  body('局部最大值（Local Maximum）：如果对包含 c 点在内的某个开区间内所有 x 值，不等式 f(x) ≤ f(c) 都成立，那么 f(c) 就是函数 f(x) 的一个局部最大值。局部最大值，有时候也称为相对最大值（Relative Maximum）。'),
  body('有局部就有全局。'),
  body('全局最大值（Global Maximum）：如果对所有的 x 值，不等式 f(x) ≤ f(c) 都成立，那么 f(c) 就是函数 f(x) 的全局最大值。全局最大值，有时候也被称作绝对最大值（Absolute Maximum）。'),
  body('接着聊一下对这两个概念的理解。所谓局部，就是只要不等式 f(x) ≤ f(c) 在哪怕很小的区间内成立，f(c) 就是局部最大值。'),
  body('例如，如果把地球上的海拔高度作为函数，那么五岳之巅的高度都是局部最大值。与此同时，你回家上楼的每一级台阶高度，甚至你放在地上的拖鞋的高度，也都是局部最大值。而珠穆朗玛峰才是这个函数的全球最大值。'),
  body('最小值的概念与最大值的概念相对。')
);

// ──────────────────────────────────────────────
// Assemble document
// ──────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: bodyFont, size: bodySize } }
    },
    paragraphStyles: [
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { font: 'SimHei', size: 28, bold: true, color: '000000' },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },           // US Letter
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }  // ~0.75 inch
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Written: ${OUT} (${(buf.length/1024).toFixed(0)} KB)`);
}).catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
