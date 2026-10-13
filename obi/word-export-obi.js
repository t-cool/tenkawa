/**
 * ポスターセッション帯活動支援資料 Wordファイル生成スクリプト
 * 
 * 要件:
 * 1. ポスターセッション・ガイド（4学年区分）と振り返り用シート（4学年区分）を１つの Word ファイルとして生成
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
     * セクション見出し（■ または 【1】等：大区分）
     */
    function createSectionHeading(text, subBadge = '') {
        return new docx.Paragraph({
            heading: docx.HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 80 },
            border: {
                bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: '94A3B8' }
            },
            children: [
                gothicRun(text, { bold: true, size: 24, color: '0F172A' }), // 12pt
                ...(subBadge ? [gothicRun('　' + subBadge, { size: 18, color: '0F766E' })] : [])
            ]
        });
    }

    /**
     * 小見出し（【発表の型】等）
     */
    function createSubheading(text) {
        return new docx.Paragraph({
            heading: docx.HeadingLevel.HEADING_3,
            spacing: { before: 140, after: 60 },
            children: [
                gothicRun(text, { bold: true, size: 21, color: '1E293B' }) // 10.5pt
            ]
        });
    }

    /**
     * 項目段落（ラベル＋説明文、インデント可）
     */
    function createItemParagraph(label, body, indentTwips = 0) {
        return new docx.Paragraph({
            spacing: { before: 20, after: 30, line: 260 },
            indent: indentTwips ? { left: indentTwips } : undefined,
            children: [
                ...(label ? [gothicRun(label, { bold: true, size: 20, color: '1E293B' })] : []),
                ...(body ? [minchoRun(body, { size: 20, color: '334155' })] : [])
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
                minchoRun(text, { size: 20, color: '1E293B' })
            ]
        });
    }

    /**
     * コツ・心得コールアウト段落（左アクセント線）
     */
    function createCalloutParagraph(label, body) {
        return new docx.Paragraph({
            spacing: { before: 100, after: 80, line: 260 },
            indent: { left: 200 },
            border: {
                left: { style: docx.BorderStyle.SINGLE, size: 18, color: '0D9488' }
            },
            children: [
                gothicRun('  ' + label + ' ', { bold: true, size: 19, color: '0F766E' }),
                minchoRun(body, { size: 19, color: '334155' })
            ]
        });
    }

    /**
     * 記入用罫線（テーブルを使わず点線下線で安全に出力）
     */
    function createInputLines(lineCount = 2) {
        const paragraphs = [];
        for (let i = 0; i < lineCount; i++) {
            paragraphs.push(new docx.Paragraph({
                spacing: { before: 60, after: 60 },
                border: {
                    bottom: { style: docx.BorderStyle.DOTTED, size: 6, color: 'CBD5E1' }
                },
                children: [
                    minchoRun('　', { size: 20 })
                ]
            }));
        }
        return paragraphs;
    }

    // ─────────────────────────────────────────────────────────────
    // 1. ポスターセッション・ガイド（1~2年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage1() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・ガイド（1・2年生）',
            '',
            '探究の視点：〜の中で（五感と発見）'
        ));

        // 1. はっぴょうのかたち
        children.push(createSectionHeading('■ はっぴょうのかたち（五感と はっけんを つたえる）'));
        children.push(createItemParagraph('・「五感のことば」：', '見たこと・音・さわったかんじを「まるで〜みたい」とたとえる'));
        children.push(createItemParagraph('・「て（ゆびさし）」：', 'ポスターの え や しゃしんを「トントン」とゆびさす'));
        children.push(createItemParagraph('・「目とこえ」：', 'きいてくださる 人（大人の方）の かおを 見て、まっすぐ こえをとどける'));

        children.push(createSubheading('【はっぴょうの じゅんばん（1分30秒・五感でみつけた３つのこと）】'));
        children.push(createItemParagraph('1. はじめ（約15秒）：', '「わたし（ぼく）が みつけた 天川村の 〇〇 です。つたえたいことは ３つ あります。」'));
        children.push(createItemParagraph('2. つたえること（約60秒・各約20秒）：', ''));
        children.push(createItemParagraph('①「１つ目は、【みため・色やかたち】です。（トントン）ここを 見てください。まるで〜みたいです。」', '', 360));
        children.push(createItemParagraph('②「２つ目は、【さわったかんじ・おと】です。（トントン）〜してみたら、〜でした。」', '', 360));
        children.push(createItemParagraph('③「３つ目は、【すきなところ】です。（トントン）ここが〜だから、だいすきです。」', '', 360));
        children.push(createItemParagraph('3. おわり（約15秒）：', '「これで はっぴょうを おわります。きいてくださり ありがとうございました！ なんでも きいてください。」'));

        // 2. しつもん＆たいわ
        children.push(createSectionHeading('■ しつもん＆たいわ（きいてくださる 人と たのしく おはなし）'));
        children.push(createSubheading('【しつもんの ヒント（大人の方からの しつもん）】'));
        children.push(createItemParagraph('・どんな かんじ？（五感）：', '「さわったとき、どんな かんじでしたか？ どんな おとがしましたか？」'));
        children.push(createItemParagraph('・どこで？（ばしょ）：', '「天川村の どこで みつけたのですか？」'));
        children.push(createItemParagraph('・どうして？（わけ・気もち）：', '「どうして そこが いちばん すきなのですか？」'));
        children.push(createItemParagraph('・見くらべ（にてる・ちがう）：', '「ほかと にてるところや ちがうところは ありますか？」'));

        children.push(createSubheading('【こたえかたの ３ステップ】'));
        children.push(createItemParagraph('1. ', '「きいてくださり、ありがとうございます！」'));
        children.push(createItemParagraph('2. ', '「ポスターの ここを 見てください。（トントン）〜だからです。」'));
        children.push(createItemParagraph('3. ', '「つたわりましたでしょうか？（にこにこ えがおで）」'));

        children.push(createCalloutParagraph('※ リラックスの コツ：', 'テストでは ありません！ きてくださった 方に「天川村のすてき」を たのしく おつたえしましょう！'));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 2. ポスターセッション・ガイド（3~4年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage2() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・ガイド（3・4年生）',
            '',
            '探究の視点：〜について（整理と要点）'
        ));

        // 1. 発表の型（観点と事実を整理して伝える）
        children.push(createSectionHeading('■ 発表の型（観点と事実を整理して伝える）'));
        children.push(createItemParagraph('1. 観点ごとに整理して話す：', '場所や工夫など、伝えるポイントを分けて順序よく話す。'));
        children.push(createItemParagraph('2. 事実と考えを分ける：', '「調べた事実（いつ・どこで・何が）」と「自分が思ったこと」を区別する。'));
        children.push(createItemParagraph('3. ポスターを指さす：', '「写真や地図のここを見てください」と言葉に合わせて指し示す。'));
        children.push(createItemParagraph('4. 相手の目を見て話す：', '手元のメモばかり見ず、笑顔で相手（大人の方）の顔を見て声を届ける。'));

        // 2. 聞き手と対話のコツ
        children.push(createSectionHeading('■ 聞き手と対話のコツ'));
        children.push(createItemParagraph('・相づちを打ちながら聞く：', '「はい」「なるほど」とうなずきながら聞き、相手の話をしっかり受け止める。'));
        children.push(createItemParagraph('・気になったところを見つける：', '「もっと知りたいこと」や「すごいと思った工夫」をメモして質問しよう。'));

        // 3. 発表と対話の型
        children.push(createSectionHeading('■ 発表と対話の型（質問を受けて、対話を深めよう）'));
        children.push(createSubheading('【発表の型（1分30秒・観点を順序立てて伝える）】'));
        children.push(createItemParagraph('1. はじめ（約15秒）：', '「私たちが調べている【テーマ：〇〇】について発表します。地域のよさや工夫を、３つのポイントから整理して伝えます。」'));
        children.push(createItemParagraph('2. 本論（約60秒・各約20秒）：', '３つの観点を順に提示（①事実・特徴、②人々の工夫、③わかった魅力）'));
        children.push(createItemParagraph('3. おわり（約15秒）：', '「これで発表を終わります。聞いてくださり、ありがとうございました！ 何かご質問や気になった点はございますか？」'));

        children.push(createSubheading('【対話のヒント（大人の方との対話・質問）】'));
        children.push(createItemParagraph('・', '「その工夫や行事は、いつ・どこで見られるのですか？」'));
        children.push(createItemParagraph('・', '「どうしてそのやり方をしているのですか？ どんな工夫があるのですか？」'));
        children.push(createItemParagraph('・', '「昔や他の地域と比べたとき、どのような違いがあったのですか？」'));

        children.push(createSubheading('【こたえ方の ３ステップ】'));
        children.push(createItemParagraph('① ', '「ご質問ありがとうございます！」'));
        children.push(createItemParagraph('② ', '「ポスターの【ここ】をご覧ください。（指さす）調べたところ、〜〜でした。」'));
        children.push(createItemParagraph('③ ', '「伝わりましたでしょうか？（笑顔で）」'));

        children.push(createCalloutParagraph('※ リラックスのコツ：', 'テストではありませんので緊張しなくても大丈夫です！ ポスターをお見せしながら、落ち着いて丁寧にお伝えしましょう！'));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 3. ポスターセッション・ガイド（5~6年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage3() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・ガイド（5・6年生）',
            '',
            '探究の視点：〜のために(準備)（根拠と伝達）'
        ));

        // 1. 発表の型（データと理由でわかりやすく伝える）
        children.push(createSectionHeading('■ 発表の型（データと理由でわかりやすく伝える）'));
        children.push(createItemParagraph('1. 理由をデータで示す：', '「〜と思います」だけでなく、ポスターの数値・グラフを指さして「なぜそう言えるか」を伝える。'));
        children.push(createItemParagraph('2. 背景や工夫も伝える：', '自然環境・産業・歴史・村の人の思いなど、理由の背景までわかりやすく話す。'));
        children.push(createItemParagraph('3. 相手の目を見て話す：', '手元のメモばかり見ず、聞き手を見て話しかける（相手７割、資料３割）。大切な言葉の前に一呼吸置く。'));
        children.push(createItemParagraph('4. 図表やグラフを指さして話す：', 'ポスターの数値や図表を指さしながら話すと、聞き手に理由や根拠がしっかり伝わります。'));

        // 2. 聞き手と対話のコツ
        children.push(createSectionHeading('■ 聞き手と対話のコツ（対話で深める）'));
        children.push(createItemParagraph('・聞き手の反応を見ながら話す：', '相づちを打ってくれているか確認しながら、ゆっくり落ち着いて話そう。'));
        children.push(createItemParagraph('・見せたい図や写真を大きく指さす：', '「ここを見てください」と声をかけながら指さすと、相手も注目しやすくなります。'));
        children.push(createItemParagraph('・質問されたら笑顔で理由を補足する：', '「ご質問ありがとうございます！」と受け止め、ポスターの資料を指しながら補足しよう。'));

        // 3. 対話の型
        children.push(createSectionHeading('■ 対話の型（質問と対話で考えを深める）'));
        children.push(createSubheading('【発表の型（1分30秒・課題からデータ、提案へ）】'));
        children.push(createItemParagraph('1. はじめ（約15秒）：', '「私たちが探究しているテーマは【〇〇】です。調査でわかったことと私たちの考えを、３つのポイントから発表します。」'));
        children.push(createItemParagraph('2. 本論（約60秒・各約20秒）：', 'データと要因の提示（①現状とデータ、②理由や背景、③比べた良さ・強み）'));
        children.push(createItemParagraph('3. おわり（約15秒）：', '「以上の理由から、天川村のために〜〜が必要だと考えました。聞いてくださり、ありがとうございました！ ご質問やご感想があればぜひお聞かせください。」'));

        children.push(createSubheading('【探究を深める対話・質問のヒント】'));
        children.push(createItemParagraph('・', '「グラフの数字がこのようになっているのには、どのような理由があるのでしょうか？」'));
        children.push(createItemParagraph('・', '「地域の方にお聞きして、一番大切だと実感されたことは何ですか？」'));
        children.push(createItemParagraph('・', '「他の地域と比べたとき、天川村ならではの良さはどのような点でしょうか？」'));

        children.push(createSubheading('【質問への答え方】'));
        children.push(createItemParagraph('① ', '「ご質問ありがとうございます！」'));
        children.push(createItemParagraph('② ', '「ポスターの【ここ】をご覧いただけますでしょうか。取材した結果、〜〜でした。」'));
        children.push(createItemParagraph('③ ', '「〇〇さん（皆さん）はどのように思われますか？（対話をつなげる）」'));

        children.push(createCalloutParagraph('※ リラックスして話そう！：', '正解・不正解を競う場ではありません。データや理由をお示ししながら、大人の方とも対話を深めていきましょう！'));

        return { properties: pageProperties, children };
    }

    // ─────────────────────────────────────────────────────────────
    // 4. ポスターセッション・ガイド（7~9年）
    // ─────────────────────────────────────────────────────────────
    function buildCardStage4() {
        const children = [];
        children.push(...createSheetHeader(
            'ポスターセッション・セッション充実ガイド（7〜9年生）',
            '',
            '探究の視点：〜のために（自然な質疑応答と対話の深化）'
        ));

        // 1. 発表後が本番！ セッションを立ち上げる3つの仕掛け
        children.push(createSectionHeading('■ 発表後が本番！ セッションを立ち上げる3つの仕掛け'));
        children.push(createItemParagraph('1. 沈黙を生まない「呼び水の問いかけ」：', '発表の終わりに「質問ありますか？」と待たず、「〜について皆さんはどのように感じられましたか？」「実はここが一番悩んだ点です」と自分から投げかけて対話を始める。'));
        children.push(createItemParagraph('2. 手元端末・資料のスタンバイ（追加の引き出し）：', '質問を受けた瞬間、取材写真・ボツ案・詳細データを手元端末でサッと見せる。相手に見えやすい角度に向ける。'));
        children.push(createItemParagraph('3. 聞き手に合わせた対話モードの切り替え：', '下級生には目線を下げてやさしく、地域の方には生活の実感を伺い、同級生には実現への課題を問いかける。'));

        // 2. 一問一答にしない！「4つの応答ステップ」
        children.push(createSectionHeading('■ 一問一答にしない！「4つの応答ステップ」'));
        children.push(createItemParagraph('Step 1【受け止め】：', '「ご質問ありがとうございます！」「鋭い視点に着目していただき、ありがとうございます」と笑顔で頷く。'));
        children.push(createItemParagraph('Step 2【背景とエピソード】：', '「はい/いいえ」だけで終わらせず、「実は取材の際に…（端末をお見せする）」と背景や裏話を添える。'));
        children.push(createItemParagraph('Step 3【パスを返す】：', '「〇〇さん（皆さん）はどのように思われますか？」「もし〜でしたら、どのように感じられますか？」と相手を対話に巻き込む。'));
        children.push(createItemParagraph('Step 4【わからない時も宝】：', '「その視点はまだ十分に調べ切れておりませんでした。大変参考になるご指摘をありがとうございます！ 今後の探究に生かします」'));

        // 3. 対話を弾ませる実践フレーズ集 ＆ 心得
        children.push(createSectionHeading('■ 対話を弾ませる実践フレーズ集 ＆ 心得'));
        children.push(createSubheading('【発表の終わりに投げる「呼び水フレーズ」】'));
        children.push(createItemParagraph('・', '「ここまでお聞きいただき、一番印象に残ったところや気になった点はございますでしょうか？」'));
        children.push(createItemParagraph('・', '「実は今、〇〇のアイデアで悩んでいます。皆さんでしたらどちらを選ばれますか？」'));
        children.push(createItemParagraph('・', '「村の未来に向けて、皆さんの普段の生活やご経験からどのように感じられますか？」'));

        children.push(createSubheading('【聞き手（相手）に応じた対話アプローチ】'));
        children.push(createItemParagraph('・下級生（1〜6年）へ：', '「写真のここ、見たことある？」「どっちが好きかな？」とやさしく共感し、笑顔で語りかける。'));
        children.push(createItemParagraph('・地域の方・大人へ：', '「昔の様子や普段の暮らしから見て、改善できる点はありますか？」と知恵をお借りする。'));
        children.push(createItemParagraph('・同級生へ：', '「実現するためのコストや手間について、別アイデアはある？」と議論を発展させる。'));

        children.push(createCalloutParagraph('※ セッションの極意（リラックスして楽しもう！）：', '発表は「きっかけ」、質疑応答が「本体」です。テストのように完璧に答える必要はありません。地域の方や大人の方と一緒に未来を考える「作戦会議」として、対話のひとときを楽しみましょう！'));

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
        children.push(createSectionHeading('【1】じぶんの はっぴょうの ふりかえり', '（めあて：五感でみつけたことを、ポスターをゆびさして たのしくつたえよう）'));
        children.push(createSubheading('じぶんの できたこと（○をつける）：'));
        children.push(createCheckItem('えがおで おおきな声で はなせた'));
        children.push(createCheckItem('えや しゃしんを ゆびさしながら はなせた'));
        children.push(createCheckItem('「まるで〜みたい！」と たとえて つたえられた'));
        children.push(createCheckItem('きいてくれた 人の しつもんに こたえられた'));
        children.push(createItemParagraph('じぶんが いちばん つたえたかったこと・みつけたこと（メモ）：', ''));
        children.push(...createInputLines(2));

        // 【2】あいての はっぴょうへの フィードバック
        children.push(createSectionHeading('【2】あいての はっぴょうへの フィードバック', '（あいてのグループへ おくるメッセージ）'));
        children.push(createSubheading('すてきだったところ（○をつける）：'));
        children.push(createCheckItem('「〜みたい！」のたとえが よくつたわった'));
        children.push(createCheckItem('ゆびさしが わかりやすくて 見やすかった'));
        children.push(createCheckItem('にこにこ えがおで たのしく はなしていた'));
        children.push(createCheckItem('しつもんに やさしく こたえてくれた'));
        children.push(createItemParagraph('あいてのグループへの ひとことメッセージ（かいてみよう）：', ''));
        children.push(...createInputLines(2));

        // 【3】つぎの はっぴょうに むけて
        children.push(createSectionHeading('【3】つぎの はっぴょうに むけて'));
        children.push(createSubheading('つぎに がんばりたいこと（○をつける）：'));
        children.push(createCheckItem('もっと おおきな こえで はなす'));
        children.push(createCheckItem('きいてくれる 人の かお（め）を みて はなす'));
        children.push(createCheckItem('ポスターを もっと たくさん ゆびさす'));
        children.push(createCheckItem('しつもんを よくきいて こたえる'));
        children.push(createItemParagraph('つぎの はっぴょうの さくせん（ひとこと）：', ''));
        children.push(...createInputLines(2));

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
        children.push(createSectionHeading('【1】自分の発表の振り返り', '（めあて：調べたこと（事実）と自分の考えを、ポスターを見せながら順序よく話そう）'));
        children.push(createSubheading('自分の振り返り（チェック）：'));
        children.push(createCheckItem('「はじめ・なか・おわり」の順序で、事実と考えを分けて話せた'));
        children.push(createCheckItem('「手はポスター、目は聞き手」を意識できた'));
        children.push(createCheckItem('写真や見出し、図を指さしながら話せた'));
        children.push(createCheckItem('質問に対して、ポスターを示して落ち着いて答えられた'));
        children.push(createItemParagraph('発表の手応え・受けた質問と回答（メモ）：', ''));
        children.push(...createInputLines(2));

        // 【2】相手の発表へのフィードバック
        children.push(createSectionHeading('【2】相手の発表へのフィードバック', '（相手グループへのアドバイス・メッセージ）'));
        children.push(createSubheading('良かったところ（チェック）：'));
        children.push(createCheckItem('くらべたことや分けたポイントが整理されていて分かりやすかった'));
        children.push(createCheckItem('ポスターの図や写真の指さし方が上手だった'));
        children.push(createCheckItem('聞き手の目を見て話しかけてくれて、聞きやすかった'));
        children.push(createCheckItem('質問にやさしく笑顔で理由を教えてくれた'));
        children.push(createItemParagraph('相手グループへの応援メッセージ・アドバイス（記入）：', ''));
        children.push(...createInputLines(2));

        // 【3】次の発表に向けて
        children.push(createSectionHeading('【3】次の発表に向けて'));
        children.push(createSubheading('次に工夫・改善したいこと（チェック）：'));
        children.push(createCheckItem('聞き手をもっと見て話す（メモを見すぎない）'));
        children.push(createCheckItem('ポスターの大切なところをタイミングよく指さす'));
        children.push(createCheckItem('「なぜそう考えたか」の理由やエピソードをくわしく話す'));
        children.push(createCheckItem('質問に答えるときに、図や写真をサッと指さす'));
        children.push(createItemParagraph('次の発表への作戦・もっと調べたいこと（記入）：', ''));
        children.push(...createInputLines(2));

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
        children.push(createSectionHeading('【1】自分の発表の振り返り', '［自己評価： A（十分できた）・B（おおむねできた）・C（次がんばる）］'));
        children.push(createItemParagraph('めあて：', '「なぜそう言えるか」ポスターのデータを見せて、理由をわかりやすく伝えよう'));
        children.push(createSubheading('観点チェック：'));
        children.push(createCheckItem('「結論→理由・データ→具体策」の筋道で論理的に説明できた'));
        children.push(createCheckItem('ポスターのグラフや数値を効果的に指さして説明できた'));
        children.push(createCheckItem('聞き手の反応を確かめながら落ち着いて対話できた'));
        children.push(createCheckItem('質問に対し、客観的な根拠や数字を示して答えられた'));
        children.push(createItemParagraph('発表の手応え・受けた質問と対話で深まったこと（メモ）：', ''));
        children.push(...createInputLines(2));

        // 【2】相手の発表へのフィードバック
        children.push(createSectionHeading('【2】相手の発表へのフィードバック', '［相手の発表評価： A（よく伝わった）・B（伝わった）・C（もう一歩）］'));
        children.push(createSubheading('良かったところ（チェック）：'));
        children.push(createCheckItem('統計グラフや現地取材などの根拠が具体的で説得力があった'));
        children.push(createCheckItem('天川村の自然や人々のくらし、歴史の背景が盛り込まれていた'));
        children.push(createCheckItem('ポスターのグラフや写真の見せ方が工夫されていて見やすかった'));
        children.push(createCheckItem('質問を笑顔で受け止め、分かりやすく理由を答えてくれた'));
        children.push(createItemParagraph('相手グループへのアドバイス・応援メッセージ（記入）：', ''));
        children.push(...createInputLines(2));

        // 【3】次の発表に向けて
        children.push(createSectionHeading('【3】次の発表に向けて'));
        children.push(createSubheading('次に高めたいポイント（チェック）：'));
        children.push(createCheckItem('データの読み取りや比較（過去・他地域との比較）を深める'));
        children.push(createCheckItem('ポスターのグラフや写真の指さしをよりスムーズにする'));
        children.push(createCheckItem('異なる立場（住民・観光客など）の視点を取り入れて提案を練る'));
        children.push(createCheckItem('質問に対し、追加データや事例をサッと出せるように準備する'));
        children.push(createItemParagraph('次の発表への改善策・探究の次の一歩（記入）：', ''));
        children.push(...createInputLines(2));

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
        children.push(createSectionHeading('【1】発表後の対話を始める「仕掛け」（事前準備・作戦）'));
        children.push(createItemParagraph('めあて：', '発表の型にとらわれず、発表後の自然な質疑応答と対話でセッションを充実させよう！'));
        children.push(createItemParagraph('💡 コツ1（呼び水）：', '「何か質問は？」と待たず、「皆さんはどのように感じられましたか？」と自分から投げかけて対話のきっかけをつくる。'));
        children.push(createItemParagraph('💡 コツ2（引き出し）：', '質問されたときにサッと見せられる取材写真、ボツ案、端末のデータなど「隠し味」を手元に準備しておく。'));
        children.push(createItemParagraph('発表の終わりに聞き手へ投げる「呼び水（問いかけの一言）」：', ''));
        children.push(...createInputLines(2));
        children.push(createItemParagraph('質問に備えて用意した「追加の引き出し（端末写真・データ・裏話）」：', ''));
        children.push(...createInputLines(2));

        // 【2】自然な質疑応答を生む「対話のキャッチボール」（実践スキル）
        children.push(createSectionHeading('【2】自然な質疑応答を生む「対話のキャッチボール」（実践スキル）', '［充実度： A（自然に対話が弾んだ）・B（丁寧に返せた）・C（一問一答）］'));
        children.push(createItemParagraph('対話の技：', '①受け止め（共感・感謝）／②背景・理由を添えて返す／③パスを返す（問い返し）／④わからない時も宝にする'));
        children.push(createCheckItem('「ご質問ありがとうございます」等のクッション言葉を使い、聞き手が話しやすい温かい空気をつくった'));
        children.push(createCheckItem('一問一答で終わらせず、「実は〜」と背景や手元端末のデータを添えて話を広げた'));
        children.push(createCheckItem('「〇〇さん（皆さん）はどう感じられますか？」と相手にもパスを返し、自然な会話の往復をつくった'));
        children.push(createCheckItem('分からない質問にも焦らず誠実に向き合い、「今後の探究で調べてみます！」と探究のタネにした'));

        // 【3】交わされた対話の記録 ＆ セッションで深まった発見
        children.push(createSectionHeading('【3】交わされた対話の記録 ＆ セッションで深まった新しい発見'));
        children.push(createItemParagraph('聞き手から出た質問・意見 ＆ 自分の応答・広がった対話（メモ）：', ''));
        children.push(...createInputLines(2));
        children.push(createItemParagraph('対話を通して新しく気づいたこと・アイデアが深まった論点（セッションの収穫）：', ''));
        children.push(...createInputLines(2));

        // 【4】次のセッションをさらに充実させるための改善 ＆ 相手へのメッセージ
        children.push(createSectionHeading('【4】次のセッションをさらに充実させるための改善 ＆ 相手へのメッセージ'));
        children.push(createSubheading('次回挑戦したい「セッション向上の工夫」（チェック）：'));
        children.push(createCheckItem('聞き手の反応を見て、呼び水となる問いかけのバリエーションを増やす'));
        children.push(createCheckItem('下級生や静かな聞き手にも「ここはどう思われますか？」など優しく話を振る'));
        children.push(createCheckItem('質問に対して、端末の写真や追加の具体例をさらにスムーズに提示する'));
        children.push(createCheckItem('相手の意見を肯定しながら「さらに〜するとどうでしょうか？」と議論を発展させる'));
        children.push(createItemParagraph('次回のセッションをもっと充実させる作戦 ／ 相手グループへのアドバイス（記入）：', ''));
        children.push(...createInputLines(2));

        return { properties: pageProperties, children };
    }

    /**
     * ドキュメント全体の生成（8つのワークシートを各セクション＝改ページとして結合）
     */
    function buildObiDocument() {
        const sections = [
            // ポスターセッション・ガイド 4種
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

            const fileName = 'ポスターセッション支援資料（ポスターセッション・ガイド・振り返りシート）.docx';
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
