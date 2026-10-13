/**
 * ポスターセッション帯活動支援資料 Wordファイル生成スクリプト
 * 
 * 要件:
 * 1. お助けカード（4学年区分）と振り返り用シート（4学年区分）を１つの Word ファイルとして生成
 * 2. それぞれのワークシートは改ページ（別セクション）される
 * 3. 見出しはゴシック体、本文は明朝体
 * 4. 上下左右の余白は 25mm (1417 twips)
 */

(function () {
    'use strict';

    const FONT_GOTHIC = '游ゴシック';
    const FONT_MINCHO = '游明朝';

    // 共通ページプロパティ（A4縦、上下10mm・左右25mm余白）
    const pageProperties = {
        page: {
            size: {
                orientation: docx.PageOrientation.PORTRAIT,
                width: docx.convertMillimetersToTwip(210), // 210mm
                height: docx.convertMillimetersToTwip(297), // 297mm
            },
            margin: {
                top: docx.convertMillimetersToTwip(10),    // 10mm = 567 twips
                bottom: docx.convertMillimetersToTwip(10), // 10mm = 567 twips
                left: docx.convertMillimetersToTwip(25),   // 25mm = 1417 twips
                right: docx.convertMillimetersToTwip(25),  // 25mm = 1417 twips
            }
        }
    };

    // 罫線スタイル
    const borderStyleThin = {
        style: docx.BorderStyle.SINGLE,
        size: 4,
        color: '94A3B8'
    };
    const bordersThin = {
        top: borderStyleThin,
        bottom: borderStyleThin,
        left: borderStyleThin,
        right: borderStyleThin
    };

    /**
     * ゴシック体テキストラン（見出し・ラベル用）
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
     * 明朝体テキストラン（本文・説明用）
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

    /**
     * シート大見出し＋サブ情報
     */
    function createSheetHeader(titleText, subtitleText, metaText = '') {
        const paragraphs = [];
        paragraphs.push(new docx.Paragraph({
            spacing: { before: 0, after: 60 },
            border: {
                bottom: { style: docx.BorderStyle.SINGLE, size: 12, color: '1E293B' }
            },
            children: [
                gothicRun(titleText, { bold: true, size: 28, color: '0F172A' }), // 14pt
                ...(subtitleText ? [
                    gothicRun('　' + subtitleText, { size: 20, color: '475569' })
                ] : [])
            ]
        }));
        if (metaText) {
            paragraphs.push(new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                spacing: { before: 40, after: 120 },
                children: [
                    gothicRun(metaText, { size: 19, color: '334155' })
                ]
            }));
        } else {
            paragraphs.push(new docx.Paragraph({ spacing: { before: 0, after: 80 } }));
        }
        return paragraphs;
    }

    /**
     * 枠付きボックス（テーブルセルによる囲み）
     */
    function createBoxTable(paragraphs, bgColor = 'F8FAFC') {
        return new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            borders: bordersThin,
                            shading: { fill: bgColor },
                            margins: { top: 100, bottom: 100, left: 140, right: 140 },
                            children: paragraphs
                        })
                    ]
                })
            ]
        });
    }

    /**
     * 解答・記入枠テーブル
     */
    function createInputBox(heightTwips = 600) {
        return new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            rows: [
                new docx.TableRow({
                    height: { value: heightTwips, rule: docx.HeightRule.ATLEAST },
                    children: [
                        new docx.TableCell({
                            borders: bordersThin,
                            shading: { fill: 'FFFFFF' },
                            margins: { top: 80, bottom: 80, left: 100, right: 100 },
                            children: [new docx.Paragraph({ text: '' })]
                        })
                    ]
                })
            ]
        });
    }

    /**
     * チェックリスト項目（□ 付き）
     */
    function createCheckItem(text) {
        return new docx.Paragraph({
            spacing: { before: 20, after: 30, line: 260 },
            indent: { left: 240, hanging: 240 },
            children: [
                gothicRun('□ ', { bold: true, size: 20, color: '1E293B' }),
                minchoRun(text, { size: 20 })
            ]
        });
    }

    // ─────────────────────────────────────────────────────────────
    // 1. お助けカード（卓上カード 1~2年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage1() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・ガイド（1・2年生）',
            '【お助けカード】',
            '探究の視点：〜の中で（五感と発見）'
        ));

        // 左ブロック：はっぴょうのかたち
        const leftContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ はっぴょうのかたち（五感と はっけんを つたえる）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30, line: 260 },
                children: [
                    gothicRun('・「五感のことば」：', { bold: true, size: 20 }),
                    minchoRun('見たこと・音・さわったかんじを「まるで〜みたい」とたとえる', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30, line: 260 },
                children: [
                    gothicRun('・「て（ゆびさし）」：', { bold: true, size: 20 }),
                    minchoRun('ポスターの え や しゃしんを「トントン」とゆびさす', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 80, line: 260 },
                children: [
                    gothicRun('・「目とこえ」：', { bold: true, size: 20 }),
                    minchoRun('おともだちの かおを 見て、まっすぐ こえをとどける', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [gothicRun('【はっぴょうの じゅんばん（1分30秒・五感でみつけた３つのこと）】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 30, after: 30, line: 260 },
                children: [
                    gothicRun('1. はじめ（約15秒）：', { bold: true, size: 20 }),
                    minchoRun('「わたし（ぼく）が みつけた 天川村の 〇〇 です。つたえたいことは ３つ あります。」', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 30, after: 30, line: 260 },
                children: [
                    gothicRun('2. つたえること（約60秒・各約20秒）：', { bold: true, size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                indent: { left: 240 },
                children: [minchoRun('①「１つ目は、【みため・色やかたち】です。（トントン）ここを 見てね。まるで〜みたいです。」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                indent: { left: 240 },
                children: [minchoRun('②「２つ目は、【さわったかんじ・おと】です。（トントン）〜してみたら、〜でした。」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 30, line: 240 },
                indent: { left: 240 },
                children: [minchoRun('③「３つ目は、【すきなところ】です。（トントン）ここが〜だから、だいすきです。」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 30, after: 30, line: 260 },
                children: [
                    gothicRun('3. おわり（約15秒）：', { bold: true, size: 20 }),
                    minchoRun('「これで はっぴょうを おわります。きいてくれて ありがとう！ なんでも きいてね。」', { size: 20 })
                ]
            })
        ];
        children.push(createBoxTable(leftContent));
        children.push(new docx.Paragraph({ spacing: { before: 60, after: 60 } }));

        // 右ブロック：しつもん＆たいわ
        const rightContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ しつもん＆たいわ（おともだちと たのしく おしゃべり）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30, line: 260 },
                children: [
                    gothicRun('・どんな かんじ？（五感）：', { bold: true, size: 20 }),
                    minchoRun('「さわったとき、どんな かんじだった？ どんな おとがした？」', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30, line: 260 },
                children: [
                    gothicRun('・どこで？（ばしょ）：', { bold: true, size: 20 }),
                    minchoRun('「天川村の どこで みつけたの？」', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30, line: 260 },
                children: [
                    gothicRun('・どうして？（わけ・気もち）：', { bold: true, size: 20 }),
                    minchoRun('「どうして そこが いちばん すきなの？」', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 60, line: 260 },
                children: [
                    gothicRun('・見くらべ（にてる・ちがう）：', { bold: true, size: 20 }),
                    minchoRun('「ぼく（わたし）のみつけたものと にてるところはあるかな？」', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [gothicRun('【こたえかたの ３ステップ】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [minchoRun('1. 「きいてくれて ありがとう！」', { size: 20 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [minchoRun('2. 「ポスターの ここを見てね。（トントン）〜だからだよ。」', { size: 20 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 60, line: 260 },
                children: [minchoRun('3. 「わかったかな？（にこにこ笑顔で）」', { size: 20 })]
            }),
            new docx.Paragraph({
                spacing: { before: 40, after: 0, line: 240 },
                children: [
                    gothicRun('※ リラックスの コツ：', { bold: true, size: 19 }),
                    minchoRun('じゅぎょうのテストじゃないよ！ おともだちと「天川村のすてき」を たのしく おしゃべりしよう！', { size: 19 })
                ]
            })
        ];
        children.push(createBoxTable(rightContent));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 2. お助けカード（卓上カード 3~4年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage2() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・ガイド（3・4年生）',
            '【お助けカード】',
            '探究の視点：〜について（整理と要点）'
        ));

        // 左ブロック：発表の型 ＆ 聞き手のコツ
        const leftContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ 発表の型（観点と事実を整理して伝える）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('1. 観点ごとに整理して話す：', { bold: true, size: 20 }),
                    minchoRun('場所や工夫など、伝えるポイントを分けて順序よく話す。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('2. 事実と考えを分ける：', { bold: true, size: 20 }),
                    minchoRun('「調べた事実（いつ・どこで・何が）」と「自分が思ったこと」を区別する。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('3. ポスターを指さす：', { bold: true, size: 20 }),
                    minchoRun('「写真や地図のここを見てください」と言葉に合わせて指し示す。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 60, line: 260 },
                children: [
                    gothicRun('4. 相手の目を見て話す：', { bold: true, size: 20 }),
                    minchoRun('手元のメモばかり見ず、笑顔で友達の顔を見て声を届ける。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [gothicRun('■ 聞き手と対話のコツ', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('・相づちを打ちながら聞く：', { bold: true, size: 20 }),
                    minchoRun('「うん、うん」「へえ〜」とうなずきながら聞き、話し手を安心させる。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('・気になったところを見つける：', { bold: true, size: 20 }),
                    minchoRun('「もっと知りたいこと」や「すごいと思った工夫」をメモして質問しよう。', { size: 20 })
                ]
            })
        ];
        children.push(createBoxTable(leftContent));
        children.push(new docx.Paragraph({ spacing: { before: 60, after: 60 } }));

        // 右ブロック：発表と対話の型
        const rightContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ 発表と対話の型（質問して、もっと仲良く知ろう）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30 },
                children: [gothicRun('【発表の型（1分30秒・観点を順序立てて伝える）】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                children: [
                    gothicRun('1. はじめ（約15秒）：', { bold: true, size: 19 }),
                    minchoRun('「私たちが調べている【テーマ：〇〇】について発表します。地域のよさや工夫を、３つのポイントから整理して伝えます。」', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                children: [
                    gothicRun('2. 本論（約60秒・各約20秒）：', { bold: true, size: 19 }),
                    minchoRun('３つの観点を順に提示（①事実・特徴、②人々の工夫、③わかった魅力）', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 40, line: 240 },
                children: [
                    gothicRun('3. おわり（約15秒）：', { bold: true, size: 19 }),
                    minchoRun('「これで発表を終わります。聞いてくれてありがとう！ 何か質問や気になったことはありますか？」', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('【対話のヒント（5W1Hで質問してみよう）】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('・「その工夫や行事は、いつ・どこで見られるの？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('・「どうしてそのやり方をしているの？ どんな工夫があるの？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 30, line: 240 },
                children: [minchoRun('・「昔や他の地域と比べたとき、どんな違いがあったの？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('【こたえ方の ３ステップ】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('①「質問ありがとう！」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('②「ポスターの【ここ】を見てね。（指さす）調べたら、〜〜だったよ。」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 40, line: 240 },
                children: [minchoRun('③「伝わったかな？（笑顔で）」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 0, line: 240 },
                children: [
                    gothicRun('※ リラックスのコツ：', { bold: true, size: 19 }),
                    minchoRun('テストじゃないから緊張しなくて大丈夫！ポスターを見せながら、普段のおしゃべりのように楽しもう！', { size: 19 })
                ]
            })
        ];
        children.push(createBoxTable(rightContent));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 3. お助けカード（卓上カード 5~6年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage3() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・ガイド（5・6年生）',
            '【お助けカード】',
            '探究の視点：〜のために(準備)（根拠と伝達）'
        ));

        // 左ブロック
        const leftContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ 発表の型（データと理由でわかりやすく伝える）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('1. 理由をデータで示す：', { bold: true, size: 20 }),
                    minchoRun('「〜と思います」だけでなく、ポスターの数値・グラフを指さして「なぜそう言えるか」を伝える。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('2. 背景や工夫も伝える：', { bold: true, size: 20 }),
                    minchoRun('自然環境・産業・歴史・村の人の思いなど、理由の背景までわかりやすく話す。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('3. 相手の目を見て話す：', { bold: true, size: 20 }),
                    minchoRun('手元のメモばかり見ず、聞き手を見て話しかける（相手７割、資料３割）。大切な言葉の前に一呼吸置く。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 60, line: 260 },
                children: [
                    gothicRun('4. 図表やグラフを指さして話す：', { bold: true, size: 20 }),
                    minchoRun('ポスターの数値や図表を指さしながら話すと、聞き手に理由や根拠がしっかり伝わります。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [gothicRun('■ 聞き手と対話のコツ（対話で深める）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('・聞き手の反応を見ながら話す：', { bold: true, size: 20 }),
                    minchoRun('相づちを打ってくれているか確認しながら、ゆっくり落ち着いて話そう。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('・見せたい図や写真を大きく指さす：', { bold: true, size: 20 }),
                    minchoRun('「ここを見てください」と声をかけながら指さすと、相手も注目しやすくなります。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('・質問されたら笑顔で理由を補足する：', { bold: true, size: 20 }),
                    minchoRun('「いい質問ありがとう！」と受け止め、ポスターの資料を指しながら補足しよう。', { size: 20 })
                ]
            })
        ];
        children.push(createBoxTable(leftContent));
        children.push(new docx.Paragraph({ spacing: { before: 60, after: 60 } }));

        // 右ブロック
        const rightContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ 対話の型（質問と対話で考えを深める）', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 30 },
                children: [gothicRun('【発表の型（1分30秒・課題からデータ、提案へ）】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                children: [
                    gothicRun('1. はじめ（約15秒）：', { bold: true, size: 19 }),
                    minchoRun('「私たちが探究しているテーマは【〇〇】です。調査でわかったことと私たちの考えを、３つのポイントから発表します。」', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                children: [
                    gothicRun('2. 本論（約60秒・各約20秒）：', { bold: true, size: 19 }),
                    minchoRun('データと要因の提示（①現状とデータ、②理由や背景、③比べた良さ・強み）', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 40, line: 240 },
                children: [
                    gothicRun('3. おわり（約15秒）：', { bold: true, size: 19 }),
                    minchoRun('「以上の理由から、天川村のために〜〜が必要だと考えました。聞いてくれてありがとう！ 質問や感想があれば教えてください。」', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('【探究を深める質問のヒント】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('・「グラフの数字がこうなっているのは、どんな理由があるのかな？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('・「地域の人に聞いてみて、自分たちで一番大切だと思ったことは何？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 30, line: 240 },
                children: [minchoRun('・「他の地域と比べたとき、天川村ならではの良さはどんなところ？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('【質問への答え方】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('①「いい質問をありがとう！」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('②「ポスターの【ここ】を見てほしいんだけど、取材した結果〜〜だったんだ。」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 40, line: 240 },
                children: [minchoRun('③「〇〇さんはどう思う？（対話をつなげる）」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 0, line: 240 },
                children: [
                    gothicRun('※ リラックスして話そう！：', { bold: true, size: 19 }),
                    minchoRun('正解・不正解を競う場ではありません。データや理由を見せながら、お互いのアイデアを楽しく語り合おう！', { size: 19 })
                ]
            })
        ];
        children.push(createBoxTable(rightContent));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 4. お助けカード（卓上カード 7~9年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage4() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・セッション充実ガイド（7〜9年生）',
            '【卓上ガイド】',
            '探究の視点：〜のために（自然な質疑応答と対話の深化）'
        ));

        // 左ブロック
        const leftContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ 発表後が本番！ セッションを立ち上げる3つの仕掛け', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('1. 沈黙を生まない「呼び水の問いかけ」：', { bold: true, size: 20 }),
                    minchoRun('発表の終わりに「質問ありますか？」と待たず、「〜について皆さんはどう感じましたか？」「実はここが一番迷った点です」と自分から投げかけて対話を始める。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20, line: 260 },
                children: [
                    gothicRun('2. 手元端末・資料のスタンバイ（追加の引き出し）：', { bold: true, size: 20 }),
                    minchoRun('質問を受けた瞬間、取材写真・ボツ案・詳細データを手元端末でサッと見せる。相手に見えやすい角度に向ける。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 40, line: 260 },
                children: [
                    gothicRun('3. 聞き手に合わせた対話モードの切り替え：', { bold: true, size: 20 }),
                    minchoRun('下級生には目線を下げてやさしく、地域の方には生活の実感を伺い、同級生には実現への課題を問いかける。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [gothicRun('■ 一問一答にしない！「4つの応答ステップ」', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 260 },
                children: [
                    gothicRun('Step 1【受け止め】：', { bold: true, size: 20 }),
                    minchoRun('「質問ありがとう！」「鋭いところに着目してくれたね」と笑顔で頷く。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 260 },
                children: [
                    gothicRun('Step 2【背景とエピソード】：', { bold: true, size: 20 }),
                    minchoRun('「はい/いいえ」だけで終わらせず、「実は取材の時に…（端末見せる）」と裏話を添える。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 260 },
                children: [
                    gothicRun('Step 3【パスを返す】：', { bold: true, size: 20 }),
                    minchoRun('「〇〇さんはどう思いますか？」「もし〜ならどうしますか？」と相手を対話に巻き込む。', { size: 20 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 260 },
                children: [
                    gothicRun('Step 4【わからない時も宝】：', { bold: true, size: 20 }),
                    minchoRun('「そこはまだ調べ切れていませんでした。良い視点をありがとう！ 一緒に考えてみませんか？」', { size: 20 })
                ]
            })
        ];
        children.push(createBoxTable(leftContent));
        children.push(new docx.Paragraph({ spacing: { before: 60, after: 60 } }));

        // 右ブロック
        const rightContent = [
            new docx.Paragraph({
                spacing: { before: 0, after: 60 },
                children: [gothicRun('■ 対話を弾ませる実践フレーズ集 ＆ 心得', { bold: true, size: 22 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('【発表の終わりに投げる「呼び水フレーズ」】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('・「ここまで聞いてみて、一番『おっ』と思ったところや気になったところはどこですか？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [minchoRun('・「実は今、〇〇のアイデアで迷っています。皆さんならどちらを選びますか？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 30, line: 240 },
                children: [minchoRun('・「村の未来に向けて、皆さんの普段の生活や体験からどう感じますか？」', { size: 19 })]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('【聞き手（相手）に応じた対話アプローチ】', { bold: true, size: 21 })]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [
                    gothicRun('・下級生（1〜6年）へ：', { bold: true, size: 19 }),
                    minchoRun('「写真のここ、見たことある？」「どっちが好きかな？」とやさしく共感し、笑顔で語りかける。', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [
                    gothicRun('・地域の方・大人へ：', { bold: true, size: 19 }),
                    minchoRun('「昔の様子や普段の暮らしから見て、改善できる点はありますか？」と知恵をお借りする。', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 30, line: 240 },
                children: [
                    gothicRun('・同級生へ：', { bold: true, size: 19 }),
                    minchoRun('「実現するためのコストや手間について、別アイデアはある？」と議論を発展させる。', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 0, line: 240 },
                children: [
                    gothicRun('※ セッションの極意（リラックスして楽しもう！）：', { bold: true, size: 19 }),
                    minchoRun('発表は「きっかけ」、質疑応答が「本体」です。テストのように完璧に答える必要はありません。聞き手と一緒に未来を考える「作戦会議」を楽しもう！', { size: 19 })
                ]
            })
        ];
        children.push(createBoxTable(rightContent));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 5. 振り返りシート（1~2年）
    // ─────────────────────────────────────────────────────────────
    function buildRecordStage1() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション ふりかえりシート（１・２年）',
            '',
            '年　組　なまえ［　　　　　　　　　］　　月　日　あいてのグループ名［　　　　　］'
        ));

        // 【1】じぶんの はっぴょうの ふりかえり
        const part1 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【1】じぶんの はっぴょうの ふりかえり', { bold: true, size: 22 }),
                    gothicRun('　（めあて：五感でみつけたことを、ポスターをゆびさして たのしくつたえよう）', { size: 19, color: '0F766E' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('じぶんの できたこと（○をつける）：', { bold: true, size: 20 })]
            }),
            createCheckItem('えがおで おおきな声で はなせた'),
            createCheckItem('えや しゃしんを ゆびさしながら はなせた'),
            createCheckItem('「まるで〜みたい！」と たとえて つたえられた'),
            createCheckItem('おともだちの しつもんに こたえられた'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('じぶんが いちばん つたえたかったこと・みつけたこと（メモ）：', { size: 19, color: '475569' })]
            }),
            createInputBox(600)
        ];
        children.push(createBoxTable(part1));
        children.push(new docx.Paragraph({ spacing: { before: 40, after: 40 } }));

        // 【2】あいての はっぴょうへの フィードバック
        const part2 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【2】あいての はっぴょうへの フィードバック', { bold: true, size: 22 }),
                    gothicRun('　（あいてのグループへ おくるメッセージ）', { size: 19, color: '475569' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('すてきだったところ（○をつける）：', { bold: true, size: 20 })]
            }),
            createCheckItem('「〜みたい！」のたとえが よくつたわった'),
            createCheckItem('ゆびさしが わかりやすくて 見やすかった'),
            createCheckItem('にこにこ えがおで たのしく はなしていた'),
            createCheckItem('しつもんに やさしく こたえてくれた'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('あいてのグループへの ひとことメッセージ（かいてみよう）：', { size: 19, color: '475569' })]
            }),
            createInputBox(600)
        ];
        children.push(createBoxTable(part2));
        children.push(new docx.Paragraph({ spacing: { before: 40, after: 40 } }));

        // 【3】つぎの はっぴょうに むけて
        const part3 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【3】つぎの はっぴょうに むけて', { bold: true, size: 22 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('つぎに がんばりたいこと（○をつける）：', { bold: true, size: 20 })]
            }),
            createCheckItem('もっと おおきな こえで はなす'),
            createCheckItem('おともだちの かお（め）を みて はなす'),
            createCheckItem('ポスターを もっと たくさん ゆびさす'),
            createCheckItem('しつもんを よくきいて こたえる'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('つぎの はっぴょうの さくせん（ひとこと）：', { size: 19, color: '475569' })]
            }),
            createInputBox(500)
        ];
        children.push(createBoxTable(part3));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 6. 振り返りシート（3~4年）
    // ─────────────────────────────────────────────────────────────
    function buildRecordStage2() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション 振り返りシート（３・４年）',
            '',
            '年　組　氏名［　　　　　　　　　］　　実施日：　月　日　相手のグループ名［　　　　　］'
        ));

        // 【1】自分の発表の振り返り
        const part1 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【1】自分の発表の振り返り', { bold: true, size: 22 }),
                    gothicRun('　（めあて：調べたこと（事実）と自分の考えを、ポスターを見せながら順序よく話そう）', { size: 19, color: '0F766E' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('自分の振り返り（チェック）：', { bold: true, size: 20 })]
            }),
            createCheckItem('「はじめ・なか・おわり」の順序で、事実と考えを分けて話せた'),
            createCheckItem('「手はポスター、目は聞き手」を意識できた'),
            createCheckItem('写真や見出し、図を指さしながら話せた'),
            createCheckItem('質問に対して、ポスターを示して落ち着いて答えられた'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('発表の手応え・受けた質問と回答（メモ）：', { size: 19, color: '475569' })]
            }),
            createInputBox(600)
        ];
        children.push(createBoxTable(part1));
        children.push(new docx.Paragraph({ spacing: { before: 40, after: 40 } }));

        // 【2】相手の発表へのフィードバック
        const part2 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【2】相手の発表へのフィードバック', { bold: true, size: 22 }),
                    gothicRun('　（相手グループへのアドバイス・メッセージ）', { size: 19, color: '475569' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('良かったところ（チェック）：', { bold: true, size: 20 })]
            }),
            createCheckItem('くらべたことや分けたポイントが整理されていて分かりやすかった'),
            createCheckItem('ポスターの図や写真の指さし方が上手だった'),
            createCheckItem('聞き手の目を見て話しかけてくれて、聞きやすかった'),
            createCheckItem('質問にやさしく笑顔で理由を教えてくれた'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('相手グループへの応援メッセージ・アドバイス（記入）：', { size: 19, color: '475569' })]
            }),
            createInputBox(600)
        ];
        children.push(createBoxTable(part2));
        children.push(new docx.Paragraph({ spacing: { before: 40, after: 40 } }));

        // 【3】次の発表に向けて
        const part3 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【3】次の発表に向けて', { bold: true, size: 22 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('次に工夫・改善したいこと（チェック）：', { bold: true, size: 20 })]
            }),
            createCheckItem('聞き手をもっと見て話す（メモを見すぎない）'),
            createCheckItem('ポスターの大切なところをタイミングよく指さす'),
            createCheckItem('「なぜそう考えたか」の理由やエピソードをくわしく話す'),
            createCheckItem('質問に答えるときに、図や写真をサッと指さす'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('次の発表への作戦・もっと調べたいこと（記入）：', { size: 19, color: '475569' })]
            }),
            createInputBox(500)
        ];
        children.push(createBoxTable(part3));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 7. 振り返りシート（5~6年）
    // ─────────────────────────────────────────────────────────────
    function buildRecordStage3() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション 振り返りシート（５・６年）',
            '',
            '年　組　氏名［　　　　　　　　　］　　実施日：　月　日　相手のグループ名［　　　　　］'
        ));

        // 【1】自分の発表の振り返り
        const part1 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【1】自分の発表の振り返り', { bold: true, size: 22 }),
                    gothicRun('　［自己評価： A（十分できた）・B（おおむねできた）・C（次がんばる）］', { size: 19, color: '0F766E' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20 },
                children: [
                    gothicRun('めあて：', { bold: true, size: 19 }),
                    minchoRun('「なぜそう言えるか」ポスターのデータを見せて、理由をわかりやすく伝えよう', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('観点チェック：', { bold: true, size: 20 })]
            }),
            createCheckItem('「結論→理由・データ→具体策」の筋道で論理的に説明できた'),
            createCheckItem('ポスターのグラフや数値を効果的に指さして説明できた'),
            createCheckItem('聞き手の反応を確かめながら落ち着いて対話できた'),
            createCheckItem('質問に対し、客観的な根拠や数字を示して答えられた'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('発表の手応え・受けた質問と対話で深まったこと（メモ）：', { size: 19, color: '475569' })]
            }),
            createInputBox(600)
        ];
        children.push(createBoxTable(part1));
        children.push(new docx.Paragraph({ spacing: { before: 40, after: 40 } }));

        // 【2】相手の発表へのフィードバック
        const part2 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【2】相手の発表へのフィードバック', { bold: true, size: 22 }),
                    gothicRun('　［相手の発表評価： A（よく伝わった）・B（伝わった）・C（もう一歩）］', { size: 19, color: '475569' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('良かったところ（チェック）：', { bold: true, size: 20 })]
            }),
            createCheckItem('統計グラフや現地取材などの根拠が具体的で説得力があった'),
            createCheckItem('天川村の自然や人々のくらし、歴史の背景が盛り込まれていた'),
            createCheckItem('ポスターのグラフや写真の見せ方が工夫されていて見やすかった'),
            createCheckItem('質問を笑顔で受け止め、分かりやすく理由を答えてくれた'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('相手グループへのアドバイス・応援メッセージ（記入）：', { size: 19, color: '475569' })]
            }),
            createInputBox(600)
        ];
        children.push(createBoxTable(part2));
        children.push(new docx.Paragraph({ spacing: { before: 40, after: 40 } }));

        // 【3】次の発表に向けて
        const part3 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    gothicRun('【3】次の発表に向けて', { bold: true, size: 22 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [gothicRun('次に高めたいポイント（チェック）：', { bold: true, size: 20 })]
            }),
            createCheckItem('データの読み取りや比較（過去・他地域との比較）を深める'),
            createCheckItem('ポスターのグラフや写真の指さしをよりスムーズにする'),
            createCheckItem('異なる立場（住民・観光客など）の視点を取り入れて提案を練る'),
            createCheckItem('質問に対し、追加データや事例をサッと出せるように準備する'),
            new docx.Paragraph({
                spacing: { before: 40, after: 20 },
                children: [gothicRun('次の発表への改善策・探究の次の一歩（記入）：', { size: 19, color: '475569' })]
            }),
            createInputBox(500)
        ];
        children.push(createBoxTable(part3));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 8. セッション充実ワークシート（7〜9年）
    // ─────────────────────────────────────────────────────────────
    function buildRecordStage4() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション セッション充実ワークシート（７〜９年）',
            '',
            '年　組　氏名［　　　　　　　　　］　　実施日：　月　日　相手［ □下級生 □同級生 □地域/大人 ］ 名前/班［　　　　　］'
        ));

        // 【1】発表後の対話を始める「仕掛け」（事前準備・作戦）
        const part1 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                    gothicRun('【1】発表後の対話を始める「仕掛け」（事前準備・作戦）', { bold: true, size: 22 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 0, after: 20 },
                children: [
                    gothicRun('めあて：', { bold: true, size: 19, color: '0F766E' }),
                    minchoRun('発表の型にとらわれず、発表後の自然な質疑応答と対話でセッションを充実させよう！', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10, line: 240 },
                children: [
                    gothicRun('💡 コツ1（呼び水）：', { bold: true, size: 19 }),
                    minchoRun('「何か質問は？」と待たず、「皆さんはどう感じましたか？」と自分から投げかけて対話のきっかけをつくる。', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                children: [
                    gothicRun('💡 コツ2（引き出し）：', { bold: true, size: 19 }),
                    minchoRun('質問されたときにサッと見せられる取材写真、ボツ案、端末のデータなど「隠し味」を手元に準備しておく。', { size: 19 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 20, after: 10 },
                children: [gothicRun('発表の終わりに聞き手へ投げる「呼び水（問いかけの一言）」：', { size: 19, color: '475569' })]
            }),
            createInputBox(350),
            new docx.Paragraph({
                spacing: { before: 20, after: 10 },
                children: [gothicRun('質問に備えて用意した「追加の引き出し（端末写真・データ・裏話）」：', { size: 19, color: '475569' })]
            }),
            createInputBox(350)
        ];
        children.push(createBoxTable(part1));
        children.push(new docx.Paragraph({ spacing: { before: 30, after: 30 } }));

        // 【2】自然な質疑応答を生む「対話のキャッチボール」（実践スキル）
        const part2 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                    gothicRun('【2】自然な質疑応答を生む「対話のキャッチボール」（実践スキル）', { bold: true, size: 22 }),
                    gothicRun('　［充実度： A（自然に対話が弾んだ）・B（丁寧に返せた）・C（一問一答）］', { size: 18, color: '0F766E' })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 20, line: 240 },
                children: [
                    gothicRun('対話の技：', { bold: true, size: 19 }),
                    minchoRun('①受け止め（共感・感謝）／②背景・理由を添えて返す／③パスを返す（問い返し）／④わからない時も宝にする', { size: 19 })
                ]
            }),
            createCheckItem('「質問ありがとう」等のクッション言葉を使い、聞き手が話しやすい温かい空気をつくった'),
            createCheckItem('一問一答で終わらせず、「実は〜」と背景や手元端末のデータを添えて話を広げた'),
            createCheckItem('「〇〇さんはどう感じますか？」と相手にもパスを返し、自然な会話の往復をつくった'),
            createCheckItem('分からない質問にも焦らず誠実に向き合い、「次調べてみます！」と探究のタネにした')
        ];
        children.push(createBoxTable(part2));
        children.push(new docx.Paragraph({ spacing: { before: 30, after: 30 } }));

        // 【3】交わされた対話の記録 ＆ セッションで深まった発見
        const part3 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                    gothicRun('【3】交わされた対話の記録 ＆ セッションで深まった新しい発見', { bold: true, size: 22 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10 },
                children: [gothicRun('聞き手から出た質問・意見 ＆ 自分の応答・広がった対話（メモ）：', { size: 19, color: '475569' })]
            }),
            createInputBox(450),
            new docx.Paragraph({
                spacing: { before: 20, after: 10 },
                children: [gothicRun('対話を通して新しく気づいたこと・アイデアが深まった論点（セッションの収穫）：', { size: 19, color: '475569' })]
            }),
            createInputBox(400)
        ];
        children.push(createBoxTable(part3));
        children.push(new docx.Paragraph({ spacing: { before: 30, after: 30 } }));

        // 【4】次へのステップ ＆ 相手へのフィードバック
        const part4 = [
            new docx.Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                    gothicRun('【4】次のセッションをさらに充実させるための改善 ＆ 相手へのメッセージ', { bold: true, size: 22 })
                ]
            }),
            new docx.Paragraph({
                spacing: { before: 10, after: 10 },
                children: [gothicRun('次回挑戦したい「セッション向上の工夫」（チェック）：', { bold: true, size: 19 })]
            }),
            createCheckItem('聞き手の反応を見て、呼び水となる問いかけのバリエーションを増やす'),
            createCheckItem('下級生や静かな聞き手にも「ここ、どう思う？」と優しく話を振る'),
            createCheckItem('質問に対して、端末の写真や追加の具体例をさらにスムーズに提示する'),
            createCheckItem('相手の意見を肯定しながら「さらに〜するとどうかな？」と議論を発展させる'),
            new docx.Paragraph({
                spacing: { before: 20, after: 10 },
                children: [gothicRun('次回のセッションをもっと充実させる作戦 ／ 相手グループへのアドバイス（記入）：', { size: 19, color: '475569' })]
            }),
            createInputBox(400)
        ];
        children.push(createBoxTable(part4));

        return { properties: pageProperties, children };
    }

    /**
     * ドキュメント全体の生成（8つのワークシートを各セクション＝改ページとして結合）
     */
    function buildObiDocument() {
        const sections = [
            // お助けカード 4種
            buildCardStage1(),
            buildCardStage2(),
            buildCardStage3(),
            buildCardStage4(),
            // 振り返りシート 4種
            buildRecordStage1(),
            buildRecordStage2(),
            buildRecordStage3(),
            buildRecordStage4()
        ];

        return new docx.Document({
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
                            size: 21,
                            color: '1E293B',
                        },
                        paragraph: {
                            spacing: { line: 280 },
                        }
                    }
                }
            },
            sections: sections
        });
    }

    /**
     * トースト通知の表示
     */
    function showToast(message, isError = false) {
        let toast = document.getElementById('obiWordExportToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'obiWordExportToast';
            toast.className = 'fixed bottom-6 right-6 z-50 px-4 py-3 rounded-lg shadow-lg text-sm font-medium flex items-center gap-2 transition-all duration-300 transform translate-y-10 opacity-0 pointer-events-none font-sans';
            document.body.appendChild(toast);
        }

        toast.className = `fixed bottom-6 right-6 z-50 px-4 py-3 rounded-lg shadow-lg text-sm font-medium flex items-center gap-2 transition-all duration-300 font-sans ${
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
    async function exportObiWordDocument() {
        if (typeof docx === 'undefined') {
            alert('Word生成ライブラリの読み込みに失敗しました。ページを再読み込みしてください。');
            return;
        }

        const btn = document.getElementById('wordExportBtn');
        const originalContent = btn ? btn.innerHTML : '';

        try {
            if (btn) {
                btn.disabled = true;
                btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin text-sm"></i><span>Word生成中...</span>';
            }

            const doc = buildObiDocument();
            const blob = await docx.Packer.toBlob(doc);

            const fileName = 'ポスターセッション支援資料（お助けカード・振り返りシート）.docx';
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            showToast('Word ファイル（全8ワークシート）をダウンロードしました');
        } catch (error) {
            console.error('Wordファイル生成エラー:', error);
            showToast('Wordファイルの生成中にエラーが発生しました', true);
            alert('Wordファイルの生成中にエラーが発生しました: ' + error.message);
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = originalContent;
            }
        }
    }

    // グローバルおよびモジュールに公開
    if (typeof window !== 'undefined') {
        window.exportObiWordDocument = exportObiWordDocument;
    }
    if (typeof module !== 'undefined' && module.exports) {
        module.exports = { buildObiDocument, exportObiWordDocument };
    }
})();
