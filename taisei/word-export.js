/**
 * 研究紀要（案） Word原稿生成スクリプト
 * 
 * 要件:
 * 1. 研究紀要（案）のデータをまとめて Word (.docx) 形式で書き出す
 * 2. 見出しはゴシック体、本文は明朝体
 * 3. 上下左右の余白は 25mm (1417 twips)
 */

(function () {
    'use strict';

    // 日本語フォント定義
    // 見出し: 游ゴシック (Yu Gothic) / MS Gothic
    // 本文: 游明朝 (Yu Mincho) / MS Mincho
    const FONT_GOTHIC = '游ゴシック';
    const FONT_MINCHO = '游明朝';

    /**
     * ゴシック体テキストラン（見出し用）
     */
    function gothicRun(text, options = {}) {
        return new docx.TextRun({
            text: text,
            font: {
                name: FONT_GOTHIC,
                ascii: 'Yu Gothic',
                eastAsia: FONT_GOTHIC,
                hAnsi: 'Yu Gothic',
                cs: 'Yu Gothic'
            },
            ...options
        });
    }

    /**
     * 明朝体テキストラン（本文用）
     */
    function minchoRun(text, options = {}) {
        return new docx.TextRun({
            text: text,
            font: {
                name: FONT_MINCHO,
                ascii: 'Yu Mincho',
                eastAsia: FONT_MINCHO,
                hAnsi: 'Yu Mincho',
                cs: 'Yu Mincho'
            },
            ...options
        });
    }

    // 紀要タイトル
    function createDocTitle(text) {
        return new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 200, after: 120 },
            children: [
                gothicRun(text, { bold: true, size: 36 }) // 18pt
            ]
        });
    }

    // 副題
    function createDocSubtitle(text) {
        return new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 40, after: 200 },
            children: [
                gothicRun(text, { bold: true, size: 24, color: '334155' }) // 12pt
            ]
        });
    }

    // 大見出し (Ⅰ, Ⅱ, Ⅲ, Ⅳ)
    function createHeading1(text) {
        return new docx.Paragraph({
            heading: docx.HeadingLevel.HEADING_1,
            spacing: { before: 360, after: 140 },
            children: [
                gothicRun(text, { bold: true, size: 28, color: '0F172A' }) // 14pt
            ]
        });
    }

    // 中見出し (1, 2, 3...)
    function createHeading2(text) {
        return new docx.Paragraph({
            heading: docx.HeadingLevel.HEADING_2,
            spacing: { before: 240, after: 100 },
            children: [
                gothicRun(text, { bold: true, size: 24, color: '1E293B' }) // 12pt
            ]
        });
    }

    // 小見出し ((1), (2)...)
    function createHeading3(text) {
        return new docx.Paragraph({
            heading: docx.HeadingLevel.HEADING_3,
            spacing: { before: 180, after: 80 },
            children: [
                gothicRun(text, { bold: true, size: 22, color: '1E293B' }) // 11pt
            ]
        });
    }

    // 細見出し (①, ②, 【第X学年】...)
    function createHeading4(text) {
        return new docx.Paragraph({
            heading: docx.HeadingLevel.HEADING_4,
            spacing: { before: 140, after: 60 },
            children: [
                gothicRun(text, { bold: true, size: 21, color: '334155' }) // 10.5pt
            ]
        });
    }

    // 本文段落
    function createParagraph(text, options = {}) {
        const isIndented = options.indent !== false;
        return new docx.Paragraph({
            spacing: { before: 40, after: 80, line: 320 },
            indent: isIndented ? { firstLine: 420 } : undefined, // 全角1字下げ (420 twips = 21pt)
            alignment: options.alignment || docx.AlignmentType.BOTH,
            children: [
                minchoRun(text, { size: 21, ...options.runOptions }) // 10.5pt
            ],
            ...options.paraOptions
        });
    }

    // リスト項目
    function createBulletItem(label, text) {
        return new docx.Paragraph({
            spacing: { before: 30, after: 60, line: 300 },
            indent: { left: 420, hanging: 420 },
            children: [
                gothicRun(label ? label + '　' : '・　', { bold: true, size: 21 }),
                minchoRun(text, { size: 21 })
            ]
        });
    }

    // キャプション
    function createCaption(text) {
        return new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 80, after: 160 },
            children: [
                gothicRun(text, { size: 19, color: '64748B' }) // 9.5pt
            ]
        });
    }

    // 学級編成と教職員数の表
    function createGakkyuTable() {
        const headers = ['学年', '１年', '２年', '３年', '４年', '５年', '６年', '７年', '８年', '９年', '特別支援', '合計', '教職員数'];
        const rowsData = [
            ['男子', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['女子', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['合計', '', '', '', '', '', '', '', '', '', '', '', '']
        ];

        const borderStyle = {
            style: docx.BorderStyle.SINGLE,
            size: 4,
            color: '94A3B8'
        };
        const borders = {
            top: borderStyle,
            bottom: borderStyle,
            left: borderStyle,
            right: borderStyle
        };

        const headerRow = new docx.TableRow({
            tableHeader: true,
            children: headers.map(h => new docx.TableCell({
                borders,
                shading: { fill: 'F1F5F9' },
                margins: { top: 100, bottom: 100, left: 60, right: 60 },
                verticalAlign: docx.VerticalAlign.CENTER,
                children: [
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [gothicRun(h, { size: 18, bold: true })]
                    })
                ]
            }))
        });

        const dataRows = rowsData.map(row => new docx.TableRow({
            children: row.map((cell, idx) => new docx.TableCell({
                borders,
                shading: idx === 0 ? { fill: 'F8FAFC' } : undefined,
                margins: { top: 100, bottom: 100, left: 60, right: 60 },
                verticalAlign: docx.VerticalAlign.CENTER,
                children: [
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [idx === 0 ? gothicRun(cell, { size: 19, bold: true }) : minchoRun(cell, { size: 19 })]
                    })
                ]
            }))
        }));

        return new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            rows: [headerRow, ...dataRows]
        });
    }

    /**
     * 画像を非同期に取得してUint8Arrayに変換
     */
    async function fetchImageBuffer(url) {
        try {
            const response = await fetch(url);
            if (!response.ok) return null;
            const buffer = await response.arrayBuffer();
            return new Uint8Array(buffer);
        } catch (e) {
            console.warn('画像の取得をスキップしました (' + url + '):', e);
            return null;
        }
    }

    /**
     * 研究紀要（案）の Word Document オブジェクトを構築
     */
    async function buildWordDocument() {
        // 画像を並列でロード
        const [soshikiImg, kousouImg, sessionImg] = await Promise.all([
            fetchImageBuffer('soshiki.png'),
            fetchImageBuffer('kousou.png'),
            fetchImageBuffer('session.jpeg')
        ]);

        const children = [];

        // ──────────────────────────────────────────────
        // ドキュメントタイトル部
        // ──────────────────────────────────────────────
        children.push(createDocTitle('ふるさと学習 研究紀要（案）'));
        children.push(createDocSubtitle('〜生活科から総合的な学習の時間へとつなぐ『ふるさと学習』の実践を通して〜'));

        // 学校情報
        children.push(new docx.Paragraph({
            alignment: docx.AlignmentType.RIGHT,
            spacing: { before: 80, after: 240 },
            children: [
                gothicRun('奈良県・天川村立天川小中学校', { size: 21, bold: true })
            ]
        }));

        // ──────────────────────────────────────────────
        // Ⅰ　学校の概要
        // ──────────────────────────────────────────────
        children.push(createHeading1('Ⅰ　学校の概要'));

        children.push(createHeading2('1　地域の特性'));
        children.push(createParagraph('本村は奈良県南部・紀伊半島の中央部に位置し、村域の４分の１が吉野熊野国立公園に指定されている。大峰山脈が連なり、近畿最高峰の八経ヶ岳をはじめ弥山や山上ヶ岳など1,500〜1,900m級の山岳が座すとともに、みたらい渓谷や名水百選に代表される清らかな水資源に恵まれている。'));
        children.push(createParagraph('古くから山岳修験の根本道場として発展し、約1300年前に役行者が開いた霊峰大峯山（山上ヶ岳）は、2004年にユネスコ世界遺産「紀伊山地の霊場と参詣道」に登録された。また、天河大弁財天社や洞川温泉など固有の歴史文化が今なお息づいている。こうした豊かな自然環境や世界遺産の歴史、文化、産業などを題材に、学校と地域が連携しながら「ふるさと学習」を推進し、地域への愛着と誇りを育む教育活動を展開している。'));

        children.push(createHeading2('2　児童生徒の実態'));
        children.push(createParagraph('本校は、令和２年に開校した県内初の施設一体型義務教育学校である。児童生徒は素直で温和、ふるさとを愛し思いやりの心をもって過ごしている。一方、少人数の親密さゆえに言葉を尽くさずとも通じ合ってしまい、筋道を立てて考えを伝え合ったり、自分の良さに気づいたりする経験が少ない。また、指示されたことには集中して取り組めるが、自ら問いを見つけて進んで行動することや、家庭で自ら学ぶ習慣の定着に課題が見られる。'));

        children.push(createHeading2('3　学校教育目標'));
        children.push(new docx.Paragraph({
            spacing: { before: 80, after: 140 },
            alignment: docx.AlignmentType.CENTER,
            children: [
                gothicRun('「ふるさとを愛し、自らの未来を主体的に切り拓く児童生徒の育成」', { bold: true, size: 23 })
            ]
        }));

        children.push(createHeading2('4　学級編成と教職員数'));
        children.push(createGakkyuTable());
        children.push(new docx.Paragraph({ spacing: { before: 60, after: 120 } }));

        // ──────────────────────────────────────────────
        // Ⅱ　研究の概要
        // ──────────────────────────────────────────────
        children.push(createHeading1('Ⅱ　研究の概要'));

        children.push(createHeading2('1　研究主題'));
        children.push(new docx.Paragraph({
            spacing: { before: 60, after: 40 },
            alignment: docx.AlignmentType.CENTER,
            children: [
                gothicRun('「９年間のつながりの中で、主体的に学び、学びを深める児童生徒の育成」', { bold: true, size: 24 })
            ]
        }));
        children.push(new docx.Paragraph({
            spacing: { before: 20, after: 140 },
            alignment: docx.AlignmentType.CENTER,
            children: [
                gothicRun('〜生活科から総合的な学習の時間へとつなぐ『ふるさと学習』の実践を通して〜', { size: 22 })
            ]
        }));

        children.push(createHeading2('2　主題設定の理由'));
        children.push(createParagraph('本主題を設定するにあたり、本校の児童生徒の実態および教育課程の状況から、大きく以下の２つの課題が挙げられた。'));
        children.push(createParagraph('１点目は、児童生徒の実態に関わる指導上の課題である。本校の児童生徒は、指示されたことに真面目に取り組む良さがある一方、受け身になりがちで、自ら問いを見出す主体性や表現力に課題が見られた。また、小規模校ゆえの親密さから阿吽の呼吸に頼り、筋道を立てて考えを伝え合う経験も不足していた。そこで、探究的な学習やポスターセッション等の対話的な活動を位置づけ、多様な意見から学び合う場を充実させることで、主体的に学ぶ態度を育成することとした。'));
        children.push(createParagraph('２点目は、ふるさと学習における教育課程上の課題である。９年間の生活科から総合的な学習の時間の中で子どもたちに身につけさせたい資質について目標設定を明確にするため、令和８年度に全体計画の見直しを行い、児童生徒自らがテーマを決定し、９年間の中で発展的に追究できる枠組みへと再構築した。低学年の生活科から後期課程の総合的な学習の時間へと学びをつなぎ、地域に根ざした自発的な探究活動を展開することで、郷土・天川村への愛着と誇りを育むこととした。'));
        children.push(createParagraph('以上のような背景から、本校では、生活科から総合的な学習の時間へとつなぐ「ふるさと学習」の実践を研究の柱に据え、全職員が一丸となって研究主題の具現化に向けた実践を進めている。'));

        children.push(createHeading2('3　研究の内容'));
        children.push(createParagraph('研究主題を達成するため、以下の２つの柱を立てて研究を進めている。', { indent: false }));
        children.push(createHeading3('(1) ９年間を通した生活科と総合的な学習の時間の実践'));
        children.push(createParagraph('施設一体型義務教育学校の強みを生かし、９年間の指導目標を全教職員で共有して系統的な学習を進めている。児童生徒自らが「知りたい・やってみたい」テーマを設定し、年度を跨いで地域を深く探究できるカリキュラムを工夫することで、主体的に課題を解決する力を育んでいる。'));
        children.push(createHeading3('(2) ポスターセッションの取組'));
        children.push(createParagraph('１年生から９年生までを通して、ポスターセッションを帯学習として日常的に実施している。ブース形式で対面しながら自分の言葉で伝え、他学年や地域の方々と双方向のやり取りを重ねることで、表現力や周りと主体的に関わる姿勢を育てている。'));

        children.push(createHeading2('4　研究組織'));
        children.push(createParagraph('校長をトップとし、研究推進委員会（校長、教頭、教務主任、研究主任）、校内研究全体会、そして３部会（天部会・木部会・川部会）および環境整備部会（校内の掲示に関わる整備を担当）からなる研究推進体制を構築している。', { indent: false }));

        if (soshikiImg) {
            children.push(new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 120, after: 60 },
                children: [
                    new docx.ImageRun({
                        data: soshikiImg,
                        transformation: { width: 520, height: 188 }
                    })
                ]
            }));
            children.push(createCaption('図：研究組織図'));
        }

        children.push(createHeading2('5　研究構想図'));
        if (kousouImg) {
            children.push(new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 120, after: 60 },
                children: [
                    new docx.ImageRun({
                        data: kousouImg,
                        transformation: { width: 480, height: 630 }
                    })
                ]
            }));
            children.push(createCaption('図：ふるさと学習 研究構想図'));
        }

        // 構想図のテキスト補足（編集・検索用）
        children.push(createHeading3('【研究構想の概要】'));
        children.push(createHeading4('［実態・教育目標・願い］'));
        children.push(createBulletItem('地域・児童生徒の実態', '地域について系統的に学ぶ機会の不足／指示待ち傾向と主体的な取組への課題／筋道を立てて考えを伝える経験の不足'));
        children.push(createBulletItem('学校教育目標', 'ふるさとを愛し、自らの未来を主体的に切り拓く児童生徒の育成'));
        children.push(createBulletItem('地域・保護者の願い', 'ふるさとへの誇りと愛着の育成／未来を生きるために必要な学力の育成／安全・安心な環境と基本的生活習慣の確立'));

        children.push(createHeading4('［めざす児童生徒像］'));
        children.push(createBulletItem('・', '自ら学び、考え、行動する子（知）'));
        children.push(createBulletItem('・', '自分と人を思いやり、ふるさと天川を愛する子（徳）'));
        children.push(createBulletItem('・', '心身ともにたくましく、夢に向かって生きぬく子（体）'));

        children.push(createHeading4('［研究の視点］'));
        children.push(createParagraph('児童生徒が主体的に学習課題に取り組むために、生活科から総合的な学習の時間へとつなぐ「ふるさと学習」を推進することで、より深い学びにつながると考える。'));

        children.push(createHeading4('［３つの実践の柱］'));
        children.push(createBulletItem('1. ふるさと学習', '児童・生徒によるテーマ設定と探究的な学び／年度を跨いで地域を深める系統的な追究'));
        children.push(createBulletItem('2. 授業中の学び合い', '話型を活用した対話的・協働的な学び／考えを深め合う日常的な対話の習慣化'));
        children.push(createBulletItem('3. ポスターセッション', 'ブース形式による対面発表と質疑応答／他学年や地域への発信・交流の実践'));

        // ──────────────────────────────────────────────
        // Ⅲ　研究の実際
        // ──────────────────────────────────────────────
        children.push(createHeading1('Ⅲ　研究の実際'));

        children.push(createHeading2('1　３部会を中心とする職員研修・授業研究の充実'));
        children.push(createHeading3('（１）前期・後期教員の交流を通じた研修'));
        children.push(createParagraph('本校の学校教育目標及び研究主題の達成を目指すために、施設一体型義務教育学校の強みを活かし、研究・研修組織として校内研究全体会の下に「天」部会・「木」部会・「川」部会という３つの部会を設けている。'));
        children.push(createParagraph('各部会には、前期課程、後期課程、教職経験年数などを考慮しながら全ての教員を配置している。各部会では、全員が授業を公開（うち１回は全体研修）し合って研究・研修を進めている。'));

        children.push(createHeading3('（２）各部会における取組'));
        children.push(createParagraph('各部会では部会ごとの重点テーマと役割を決めて、研究を進めている。'));
        children.push(createParagraph('「天」部会では、生活科から総合的な学習の時間への９年間の系統性を見通した全体計画の策定を中心に授業改善を推進している。'));
        children.push(createParagraph('「木」部会では、児童生徒の問いの深まりを支える過去の探究テーマの集積・分析や、日常的な表現力・対話力を培うポスターセッション帯活動の教材・話型支援の開発を中心に授業改善を推進している。'));
        children.push(createParagraph('「川」部会では、児童生徒の探究成果を校内掲示やデジタルアーカイブとして蓄積する「ポスターの保管」の構築や、地域住民・事業所と協働する「人材バンク」の整備・活用を中心に授業改善を推進している。'));
        children.push(createParagraph('「環境整備部会」では、主に校内の掲示に関わる整備を担当することを主な役割とし、児童生徒の探究の成果や学びの足跡を効果的に共有・発信できる教育環境の充実に努めている。'));

        children.push(createHeading3('（３）PBLに関する研修'));
        children.push(createParagraph('実社会に即した「正解のない課題」を発見し、解決策の立案から実行・発信までを自ら行う「PBL（Project-Based Learning：プロジェクト型学習）」をテーマに職員研修を実施している。これは、子どもたちの主体性を育成する有効な手法として着目するだけでなく、教員自身がふるさと学習において児童生徒の探究をファシリテートしていくための道標とすることをねらいとしている。ふるさと学習を軸に、前期課程では天川村についての学びを深め、後期課程では村への提言を行うプロジェクトへと展開させることで、学校教育の成果が広く地域へ還元される仕組みづくりを目指している。'));
        children.push(createParagraph('このPBLの手法に関する研修は、令和8年度より本格的に開始した。同年度には、PBLの実践に長年携わっている吉田正志教諭を招聘し、総合的な学習の時間において地域と学校が協働するための手立てや工夫を学ぶ研修会を開催した。あわせて、他府県の実践事例や関連書籍を教職員全体で共有し、各自が知見を深められる自学自習の環境を整えた。'));
        children.push(createParagraph('続く令和9年度には、ふるさと学習を先進的に進める他校との交流を中心とした研修を展開した。他校との情報交換や事例共有を通じて多様な知見を取り入れ、本校のふるさと学習が単発の活動にとどまらず、地域と学校の協働的な取り組みとして定着するよう体制づくりを進めた。'));
        children.push(createParagraph('こうした研修の積み重ねを通じて、学校と地域が有機的につながる教育活動が活発に展開されている。児童生徒が地域住民や関係機関と直接関わり、天川村の魅力発信や課題解決に向けた実践に取り組むなど、PBLを基盤とした地域協働の学びが着実に広がっている。'));

        children.push(createHeading3('（４）校内研究の成果と課題'));
        children.push(createParagraph('（※令和９年度の３学期に実践の検証を踏まえて原稿を作成して仕上げる。）', { runOptions: { color: '64748B' } }));

        children.push(createHeading2('2　９年間を通した「ふるさと学習」の実践'));
        children.push(createHeading3('（１）９年間を見通した探究的な学びを支える全体計画の作成'));
        children.push(createParagraph('１・２年生の生活科から３〜９年生の総合的な学習の時間へと滑らかにつながる全体計画を作成した。「知る・学ぶ（1〜2年）」「学ぶ・広める・深める（3〜4年）」「考える・深める・発信・提案する（5〜6年）」「探究の再確認・地域課題への参画と発信（7〜9年）」という学びの発展段階を設定し、９年間を見通した系統的な指導目標や育成を目指す資質・能力（知識及び技能、思考・判断・表現、主体的に学習に取り組む態度）を位置づけている。'));

        children.push(createHeading3('（２）地域社会と連携した探究的な学習環境の整備'));
        children.push(createHeading4('①「ポスターの制作、掲示および管理」による成果物の蓄積と発信'));
        children.push(createParagraph('児童生徒が作成したポスターやデジタルスライド等の成果物を「ポスターの保管」として蓄積し、校内掲示やWebアーカイブとして公開することで、次年度以降の探究の教材として活用できるようにしている。'));
        children.push(createHeading4('②「ふるさと人材バンク」の活用による地域との協働'));
        children.push(createParagraph('天川村の産業、自然、歴史、文化等に関わる村民・事業所・団体の協力情報を「ふるさと人材バンク」として整理し、ゲストティーチャーや現地インタビュー、フィールドワーク等で積極的に連携・協働している。'));

        children.push(createHeading3('（３）各学年のテーマと探究実践の実際'));
        children.push(createParagraph('各学年において、児童生徒が自らの興味・関心に基づき天川村に関わる問いを設定し、１年間を通した探究活動を展開している。'));

        const grades = [
            { grade: '【第１学年】生活科', note: '※令和９年度の探究テーマ・活動内容・児童の変容を記載する。' },
            { grade: '【第２学年】生活科', note: '※令和９年度の探究テーマ・活動内容・児童の変容を記載する。' },
            { grade: '【第３学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・児童生徒の変容を記載する。' },
            { grade: '【第４学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・児童生徒の変容を記載する。' },
            { grade: '【第５学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・児童生徒の変容を記載する。' },
            { grade: '【第６学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・児童生徒の変容を記載する。' },
            { grade: '【第７学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・生徒の変容を記載する。' },
            { grade: '【第８学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・生徒の変容を記載する。' },
            { grade: '【第９学年】総合的な学習の時間', note: '※令和９年度の探究テーマ・活動内容・生徒の変容を記載する。' },
        ];
        for (const g of grades) {
            children.push(createHeading4(g.grade));
            children.push(createParagraph(g.note, { runOptions: { color: '64748B' } }));
        }

        children.push(createHeading3('（４）各教科等・学校行事等と関連づけたふるさと学習'));
        children.push(createParagraph('生活科・総合的な学習の時間を核としつつ、各教科（社会科、理科、国語科等）や学校行事（みたらい渓谷散策、洞川散策、遠足、校外学習、防災学習等）との教科等横断的な連携を図り、体験的な学びをふるさと学習の探究へと有機的につなげている。'));

        children.push(createHeading2('3　ポスターセッションの実践'));
        children.push(createHeading3('（１）帯活動の位置づけと全校指導体制'));
        children.push(createParagraph('小規模校ゆえの課題である「阿吽の呼吸に頼ったコミュニケーション」や「受け身の姿勢」を克服し、自らの言葉で考えを筋道立てて相手に伝える力を育成するため、全校で日常的な「帯活動」を位置づけている。帯活動では、全教職員の共通理解のもと、短時間で反復して対話的な実践を重ねる指導体制を整えている。'));

        children.push(createHeading3('（２）対話的な学びを支える指導ツールの工夫'));
        children.push(createHeading4('① 学年段階別「ポスターセッション・ガイド（話型支援）」の活用'));
        children.push(createParagraph('発表者用・質問者用の話型（発表の切り出し、理由の提示、質問・感想の述べ方等）をまとめたポスターセッション・ガイドを、「低学年（1〜2年）」「中学年（3〜4年）」「高学年（5〜6年）」「後期課程（7〜9年）」の4段階で作成し、机上に提示することで、児童生徒が自発的に対話を深められるよう工夫している。'));
        children.push(createHeading4('② 自己の変容を見取る「振り返りシート」の活用'));
        children.push(createParagraph('毎回のセッション後に、他者の発表から気づいたことや自己の課題を記録する「振り返りシート」を位置づけ、発表して終わりにするのではなく、次の探究への問い直しや自己変容の自覚へとつなげている。'));

        children.push(createHeading3('（３）学習交流会の実施'));
        children.push(createParagraph('探究の成果を発表し学びを深め合う場として「学習交流会」を実施している。学習交流会では、異学年による交流会を実施し、ブース形式によるポスターセッション等を取り入れることで、全児童生徒が発表や質疑応答に参加する機会を設けている。異学年で交流することにより、下級生は上級生の発表や表現方法を参照して学び（他者参照）、上級生は相手意識をもったわかりやすい説明を心がけるなど、施設一体型義務教育学校の強みを生かした主体的な学び合いの場となっている。'));

        if (sessionImg) {
            children.push(new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 120, after: 60 },
                children: [
                    new docx.ImageRun({
                        data: sessionImg,
                        transformation: { width: 500, height: 304 }
                    })
                ]
            }));
            children.push(createCaption('図：学習交流会の様子（イメージ）'));
        }

        // ──────────────────────────────────────────────
        // Ⅳ　研究の成果と今後の課題
        // ──────────────────────────────────────────────
        children.push(createHeading1('Ⅳ　研究の成果と今後の課題'));
        children.push(createParagraph('（※令和10年度1学期に原稿を作成して仕上げる。）', { runOptions: { color: '64748B' } }));

        // ──────────────────────────────────────────────
        // Word Document 生成設定
        // ──────────────────────────────────────────────
        const doc = new docx.Document({
            styles: {
                default: {
                    document: {
                        run: {
                            font: {
                                name: FONT_MINCHO,
                                ascii: 'Yu Mincho',
                                eastAsia: FONT_MINCHO,
                                hAnsi: 'Yu Mincho',
                                cs: 'Yu Mincho'
                            },
                            size: 21, // 10.5pt (本文標準)
                            color: '1E293B',
                        },
                        paragraph: {
                            spacing: { line: 320 },
                        }
                    }
                },
                paragraphStyles: [
                    {
                        id: 'Heading1',
                        name: 'Heading 1',
                        basedOn: 'Normal',
                        next: 'Normal',
                        quickFormat: true,
                        run: {
                            font: { name: FONT_GOTHIC, ascii: 'Yu Gothic', eastAsia: FONT_GOTHIC, hAnsi: 'Yu Gothic', cs: 'Yu Gothic' },
                            size: 28, // 14pt
                            bold: true,
                            color: '0F172A',
                        },
                        paragraph: {
                            spacing: { before: 360, after: 140 },
                        }
                    },
                    {
                        id: 'Heading2',
                        name: 'Heading 2',
                        basedOn: 'Normal',
                        next: 'Normal',
                        quickFormat: true,
                        run: {
                            font: { name: FONT_GOTHIC, ascii: 'Yu Gothic', eastAsia: FONT_GOTHIC, hAnsi: 'Yu Gothic', cs: 'Yu Gothic' },
                            size: 24, // 12pt
                            bold: true,
                            color: '1E293B',
                        },
                        paragraph: {
                            spacing: { before: 240, after: 100 },
                        }
                    },
                    {
                        id: 'Heading3',
                        name: 'Heading 3',
                        basedOn: 'Normal',
                        next: 'Normal',
                        quickFormat: true,
                        run: {
                            font: { name: FONT_GOTHIC, ascii: 'Yu Gothic', eastAsia: FONT_GOTHIC, hAnsi: 'Yu Gothic', cs: 'Yu Gothic' },
                            size: 22, // 11pt
                            bold: true,
                            color: '1E293B',
                        },
                        paragraph: {
                            spacing: { before: 180, after: 80 },
                        }
                    },
                    {
                        id: 'Heading4',
                        name: 'Heading 4',
                        basedOn: 'Normal',
                        next: 'Normal',
                        quickFormat: true,
                        run: {
                            font: { name: FONT_GOTHIC, ascii: 'Yu Gothic', eastAsia: FONT_GOTHIC, hAnsi: 'Yu Gothic', cs: 'Yu Gothic' },
                            size: 21, // 10.5pt
                            bold: true,
                            color: '334155',
                        },
                        paragraph: {
                            spacing: { before: 140, after: 60 },
                        }
                    }
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size: {
                            orientation: docx.PageOrientation.PORTRAIT,
                            width: docx.convertMillimetersToTwip(210), // A4 幅 210mm
                            height: docx.convertMillimetersToTwip(297), // A4 高さ 297mm
                        },
                        margin: {
                            // 上下左右 25mm 指定 (1417 twips)
                            top: docx.convertMillimetersToTwip(25),
                            bottom: docx.convertMillimetersToTwip(25),
                            left: docx.convertMillimetersToTwip(25),
                            right: docx.convertMillimetersToTwip(25),
                        }
                    }
                },
                children: children
            }]
        });

        return doc;
    }

    /**
     * トースト通知の表示
     */
    function showToast(message, isError = false) {
        let toast = document.getElementById('wordExportToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'wordExportToast';
            toast.className = 'fixed bottom-6 right-6 z-50 px-4 py-3 rounded-lg shadow-lg text-sm font-medium flex items-center gap-2 transition-all duration-300 transform translate-y-10 opacity-0 pointer-events-none font-gothic';
            document.body.appendChild(toast);
        }

        toast.className = `fixed bottom-6 right-6 z-50 px-4 py-3 rounded-lg shadow-lg text-sm font-medium flex items-center gap-2 transition-all duration-300 font-gothic ${
            isError ? 'bg-red-600 text-white' : 'bg-slate-900 text-white'
        }`;
        toast.innerHTML = isError
            ? `<i class="fa-solid fa-circle-exclamation text-base"></i><span>${message}</span>`
            : `<i class="fa-solid fa-circle-check text-emerald-400 text-base"></i><span>${message}</span>`;

        const raf = (typeof window !== 'undefined' && window.requestAnimationFrame) ? window.requestAnimationFrame : function(cb) { setTimeout(cb, 16); };
        raf(() => {
            toast.classList.remove('translate-y-10', 'opacity-0');
            toast.classList.add('translate-y-0', 'opacity-100');
        });

        setTimeout(() => {
            toast.classList.remove('translate-y-0', 'opacity-100');
            toast.classList.add('translate-y-10', 'opacity-0');
        }, 4000);
    }

    /**
     * Word エクスポート実行関数 (ボタンクリック時に呼ばれる)
     */
    async function exportWordDocument() {
        if (typeof docx === 'undefined') {
            alert('Word生成ライブラリの読み込みに失敗しました。ページを再読み込みしてください。');
            return;
        }

        const btn = document.getElementById('wordExportBtn');
        const originalContent = btn ? btn.innerHTML : '';

        try {
            if (btn) {
                btn.disabled = true;
                btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin text-xs"></i><span>生成中...</span>';
            }

            // ドキュメント構築
            const doc = await buildWordDocument();

            // Blob生成
            const blob = await docx.Packer.toBlob(doc);

            // ダウンロードトリガー
            const fileName = 'ふるさと学習_研究紀要（案）.docx';
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            showToast('Word 原稿（.docx）をダウンロードしました');
        } catch (error) {
            console.error('Word原稿生成エラー:', error);
            showToast('Word原稿の生成中にエラーが発生しました', true);
            alert('Word原稿の生成中にエラーが発生しました: ' + error.message);
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = originalContent;
            }
        }
    }

    // グローバルにエクスポート
    window.exportWordDocument = exportWordDocument;
})();
