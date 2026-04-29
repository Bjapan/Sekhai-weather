const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function heading(text, color = "185FA5") {
    return new Paragraph({
        spacing: { before: 280, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color } },
        children: [new TextRun({ text, bold: true, size: 26, color, font: "Arial" })]
    });
}

function label(text) {
    return new Paragraph({
        spacing: { before: 160, after: 60 },
        children: [new TextRun({ text, bold: true, size: 20, color: "444444", font: "Arial" })]
    });
}

function body(text, color = "222222") {
    return new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text, size: 20, color, font: "Arial" })]
    });
}

function note(text) {
    return new Paragraph({
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text, size: 18, color: "888888", italics: true, font: "Arial" })]
    });
}

function box(lines, bgColor = "FFF3E0", textColor = "7B4F00") {
    return new Table({
        width: { size: 8600, type: WidthType.DXA },
        columnWidths: [8600],
        rows: [new TableRow({
            children: [new TableCell({
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 2, color: "BA7517" },
                    bottom: { style: BorderStyle.SINGLE, size: 2, color: "BA7517" },
                    left: { style: BorderStyle.SINGLE, size: 10, color: "BA7517" },
                    right: { style: BorderStyle.SINGLE, size: 2, color: "BA7517" },
                },
                shading: { fill: bgColor, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 120 },
                children: lines.map(l => new Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [new TextRun({ text: l, size: 19, color: textColor, font: "Arial" })]
                }))
            })]
        })]
    });
}

function greenBox(lines) {
    return new Table({
        width: { size: 8600, type: WidthType.DXA },
        columnWidths: [8600],
        rows: [new TableRow({
            children: [new TableCell({
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 2, color: "1D9E75" },
                    bottom: { style: BorderStyle.SINGLE, size: 2, color: "1D9E75" },
                    left: { style: BorderStyle.SINGLE, size: 10, color: "1D9E75" },
                    right: { style: BorderStyle.SINGLE, size: 2, color: "1D9E75" },
                },
                shading: { fill: "E8F8F2", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 120 },
                children: lines.map(l => new Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [new TextRun({ text: l, size: 19, color: "085041", font: "Arial" })]
                }))
            })]
        })]
    });
}

function space(n = 100) {
    return new Paragraph({ spacing: { before: n, after: n }, children: [new TextRun("")] });
}

const doc = new Document({
    styles: {
        default: { document: { run: { font: "Arial", size: 20 } } }
    },
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 1000, right: 900, bottom: 1000, left: 900 }
            }
        },
        children: [

            // TITLE
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 200 },
                children: [new TextRun({ text: "BISHAL PULAMI MAGAR — CV FIX SHEET", bold: true, size: 30, color: "185FA5", font: "Arial" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 60 },
                children: [new TextRun({ text: "Copy the fixed text below into your school CV form | 以下のテキストをCVに貼り付けてください", size: 18, color: "888888", font: "Arial" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 320 },
                children: [new TextRun({ text: "Total time to finish: 2-3 hours  |  今日中に完成できます", size: 18, color: "E24B4A", bold: true, font: "Arial" })]
            }),

            // SCHEDULE
            heading("⏰  Today's Schedule — Complete in order", "E24B4A"),
            space(60),
            box([
                "Step 1  (30 min)  — Open your school CV template. Fix the typo: 「館jきょう」→「環境」",
                "Step 2  (45 min)  — Delete old 志望動機. Type in the new text from Section A below.",
                "Step 3  (45 min)  — Delete old 自己PR. Type in the new text from Section B below.",
                "Step 4  (30 min)  — Read everything aloud once. Check for any typos. Save & print.",
                "Total: about 2.5 hours  |  You can finish tonight.",
            ]),
            space(120),

            // TYPO FIX
            heading("🔴  URGENT — Fix this typo first"),
            space(60),
            new Table({
                width: { size: 8600, type: WidthType.DXA },
                columnWidths: [4300, 4300],
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                borders,
                                shading: { fill: "FCEBEB", type: ShadingType.CLEAR },
                                margins: { top: 100, bottom: 100, left: 140, right: 140 },
                                children: [
                                    new Paragraph({ children: [new TextRun({ text: "WRONG (delete this)", bold: true, size: 19, color: "A32D2D", font: "Arial" })] }),
                                    new Paragraph({ children: [new TextRun({ text: "館jきょう", size: 22, color: "A32D2D", font: "Arial" })] }),
                                ]
                            }),
                            new TableCell({
                                borders,
                                shading: { fill: "E1F5EE", type: ShadingType.CLEAR },
                                margins: { top: 100, bottom: 100, left: 140, right: 140 },
                                children: [
                                    new Paragraph({ children: [new TextRun({ text: "CORRECT (use this)", bold: true, size: 19, color: "085041", font: "Arial" })] }),
                                    new Paragraph({ children: [new TextRun({ text: "環境 (かんきょう)", size: 22, color: "085041", font: "Arial" })] }),
                                ]
                            }),
                        ]
                    })
                ]
            }),
            space(120),

            // SECTION A - SHIBO DOKI
            heading("Section A — 志望動機（新しいバージョン）"),
            space(60),
            label("↓ Type this into your CV form exactly as written below ↓"),
            space(80),
            greenBox([
                "データとAIを活用した日本のDX推進に貢献したいと考え、ジールを志望しました。",
                "",
                "貴社は、コンサルティングからデータ基盤の構築・運用保守まで一気通貫で支援する体制と、",
                "年間350件を超える豊富な実績を誇っており、データ活用のプロフェッショナルとして",
                "成長できる最高の環境だと確信しています。",
                "",
                "現在、日本国際工科専門学校のITマネジメントコースにて、システム設計と",
                "データ活用の基礎を学んでいます。また、HTML・CSSなどのプログラミングスキルに加え、",
                "生成AIを活用したプロンプトエンジニアリングを独学で習得しており、",
                "貴社が推進するAI-Readyな基盤づくりへの関心は非常に高いです。",
                "",
                "将来はデータエンジニアとしての技術を磨き、お客様のDX推進を支える",
                "コンサルタントへと成長することを目標としています。",
                "貴社の一気通貫のサービス体制のもとで、長期的に貢献したいと考えています。",
            ]),
            space(100),
            label("English translation (for your reference only — do NOT put this in the CV)"),
            body("I am applying to Zeal because I want to contribute to Japan's DX promotion through data and AI.", "666666"),
            body("Zeal's end-to-end support model — from consulting through data platform construction to", "666666"),
            body("maintenance — combined with 350+ annual projects, makes it the ideal environment to grow as", "666666"),
            body("a data professional. I am studying IT Management at Japan International Institute of Cybernetics,", "666666"),
            body("learning system design and data utilization. I have also independently learned prompt engineering", "666666"),
            body("using generative AI, which aligns strongly with Zeal's AI-Ready direction. My goal is to grow", "666666"),
            body("from a data engineer into a DX consultant, contributing long-term within Zeal's service model.", "666666"),
            space(160),

            // SECTION B - JIKO PR
            heading("Section B — 自己PR（新しいバージョン）"),
            space(60),
            label("↓ Type this into your CV form exactly as written below ↓"),
            space(80),
            greenBox([
                "私の強みは、目標に向かって粘り強く学び続ける継続力と、多文化環境で培った高い適応力です。",
                "",
                "ネパールでコンピューターサイエンスのディプロマを取得後、さらなる成長を目指して来日し、",
                "現在は日本国際工科専門学校のITマネジメントコースでシステム設計・データ活用の",
                "基礎を学んでいます。毎日プログラミングと日本語の学習を欠かさず続けており、",
                "継続的な自己成長を習慣にしています。",
                "",
                "技術面では、HTML・CSS・Word・Excelに加え、生成AIを活用した",
                "プロンプトエンジニアリングを独学で習得しました。AIとデータの融合という",
                "最新トレンドへの関心が高く、現在PythonとSQLの学習にも取り組んでいます。",
                "",
                "また、調理・配膳・品出し・接客と4種類のアルバイトを経験したことで、",
                "どんな環境でもチームに貢献できる柔軟性と責任感を培いました。",
                "貴社でジュニアエンジニアとして着実に技術を積み、",
                "将来はデータ活用のプロフェッショナルとして長期的に貢献したいと考えています。",
            ]),
            space(100),
            label("English translation (for your reference only — do NOT put this in the CV)"),
            body("My strengths are persistent learning toward clear goals and the adaptability built through", "666666"),
            body("multicultural experience. After earning a Computer Science Diploma in Nepal, I came to Japan", "666666"),
            body("to grow further and am studying IT Management, studying both programming and Japanese every", "666666"),
            body("single day. Beyond HTML, CSS, Word and Excel, I independently learned prompt engineering", "666666"),
            body("using generative AI — directly connecting to the AI and data convergence trend. I am now", "666666"),
            body("learning Python and SQL. My four part-time roles built genuine flexibility and reliability.", "666666"),
            body("I am committed to growing steadily as a junior engineer and contributing long-term at Zeal.", "666666"),
            space(160),

            // CHECKLIST
            heading("✅  Final Checklist Before Submitting", "1D9E75"),
            space(60),
            box([
                "□  Typo fixed: 「環境」(not 館jきょう)",
                "□  志望動機 mentions: ジール / DX / データ / AI-Ready / 一気通貫 / 350件",
                "□  自己PR mentions: プロンプトエンジニアリング / ITマネジメントコース / 毎日勉強",
                "□  No generic phrases like 「どの会社でもよい」or 「チームワークが得意」alone",
                "□  Read the whole CV out loud once — your ear catches mistakes your eyes miss",
                "□  Ask your school's Japanese teacher to check the Japanese text if possible",
                "□  Print one test copy and check the layout before the final print",
            ], "FFF8E1", "5D4037"),
            space(120),

            // NEXT STEPS
            heading("📅  After Submitting — What to do next", "185FA5"),
            space(60),
            new Table({
                width: { size: 8600, type: WidthType.DXA },
                columnWidths: [1600, 7000],
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ borders, shading: { fill: "E6F1FB", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "This week", bold: true, size: 19, color: "0C447C", font: "Arial" })] })] }),
                            new TableCell({ borders, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Register for JLPT N2 (July 2026 exam). Check visa renewal at your school's office.", size: 19, color: "222222", font: "Arial" })] })] }),
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ borders, shading: { fill: "E6F1FB", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "1 month", bold: true, size: 19, color: "0C447C", font: "Arial" })] })] }),
                            new TableCell({ borders, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Start Python basics (Kaggle free course). Start SQL basics (SQLZoo.net — free).", size: 19, color: "222222", font: "Arial" })] })] }),
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ borders, shading: { fill: "E6F1FB", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "3 months", bold: true, size: 19, color: "0C447C", font: "Arial" })] })] }),
                            new TableCell({ borders, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Attend one FREE Zeal online seminar. Build one small Python data project on GitHub.", size: 19, color: "222222", font: "Arial" })] })] }),
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ borders, shading: { fill: "E6F1FB", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "2027", bold: true, size: 19, color: "0C447C", font: "Arial" })] })] }),
                            new TableCell({ borders, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Apply for Zeal 新卒採用 2027. With N2 + Python/SQL + project, you are a real candidate.", size: 19, color: "222222", font: "Arial" })] })] }),
                        ]
                    }),
                ]
            }),

            space(200),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "You can do this, Bishal. Start with Step 1 right now.", size: 22, color: "185FA5", bold: true, font: "Arial" })]
            }),
        ]
    }]
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync('/home/claude/bishal_cv_fix.docx', buf);
    console.log('Done');
});