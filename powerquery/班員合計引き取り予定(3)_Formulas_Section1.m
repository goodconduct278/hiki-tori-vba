section Section1;

shared fnNormalize = (s as text) as text =>
let
    narrow   = Text.Lower(Text.From(s)),
    noSpace  = Text.Remove(narrow, {" ", "　", Character.FromNumber(160), Character.FromNumber(9)}),
    noDot    = Text.Remove(noSpace, {".", "．", "・"}),
    hyphen1  = Text.Replace(noDot,   "－", "-"),
    hyphen2  = Text.Replace(hyphen1, "―", "-"),
    hyphen3  = Text.Replace(hyphen2, "ｰ",  "-"),
    slash    = Text.Replace(hyphen3, "／", "/"),
    paren1   = Text.Replace(slash,  "（", "("),
    paren2   = Text.Replace(paren1, "）", ")"),
    paren3   = Text.Replace(paren2, "［", "["),
    paren4   = Text.Replace(paren3, "］", "]"),
    paren5   = Text.Replace(paren4, "｛", "{"),
    paren6   = Text.Replace(paren5, "｝", "}"),
    ml1      = Text.Replace(paren6, "ｍｌ", "ml"),
    ml2      = Text.Replace(ml1,    "㎖",   "ml"),
    l1       = Text.Replace(ml2,    "ｌ",   "l"),
    result   = Text.Trim(l1)
in
    result;

shared 変換リスト = let
    ソース = Excel.CurrentWorkbook(){[Name="テーブル1"]}[Content],
    変更された型 = Table.TransformColumnTypes(ソース,{{"変換前（材料名）", type text}, {"変換後（製品名）", type text}, {"UR", type text}, {"メーカー", type text}})
in
    変更された型;

shared #"★班員データ" = let
    フォルダ設定 = Excel.CurrentWorkbook(){[Name="班員データフォルダ"]}[Content],
    FolderPath = Text.From(フォルダ設定{0}[Column1]),

    ソース = Folder.Files(FolderPath),
    非表示ファイルのフィルタ = Table.SelectRows(ソース, each [Attributes]?[Hidden]? <> true),
    必要な列の選択 = Table.SelectColumns(非表示ファイルのフィルタ, {"Name", "Content"}),
    ブックの展開 = Table.AddColumn(必要な列の選択, "Data", each Excel.Workbook([Content], null, true)),
    テーブルの展開 = Table.ExpandTableColumn(ブックの展開, "Data", {"Name", "Data", "Item", "Kind"}, {"Sheet.Name", "Sheet.Data", "Item", "Kind"}),
    集計テーブルのフィルタ = Table.SelectRows(テーブルの展開, each ([Item] = "集計テーブル" and [Kind] = "Table")),

    標準化Data追加 = Table.AddColumn(
        集計テーブルのフィルタ,
        "標準化Data",
        each
            Table.SelectColumns(
                [Sheet.Data],
                {
                    "担当者",
                    "得意先",
                    "現場",
                    "材料",
                    "数量",
                    "単位",
                    "納品日",
                    "注文状況",
                    "注文日",
                    "現場状況",
                    "チェック",
                    "日時",
                    "UR"
                },
                MissingField.UseNull
            )
    ),

    不要な列の削除 = Table.SelectColumns(標準化Data追加, {"Name", "標準化Data"}),

    データの展開 = Table.ExpandTableColumn(
        不要な列の削除,
        "標準化Data",
        {
            "担当者",
            "得意先",
            "現場",
            "材料",
            "数量",
            "単位",
            "納品日",
            "注文状況",
            "注文日",
            "現場状況",
            "チェック",
            "日時",
            "UR"
        },
        {
            "担当者",
            "得意先",
            "現場",
            "材料",
            "数量",
            "単位",
            "納品日",
            "注文状況",
            "注文日",
            "現場状況",
            "チェック",
            "日時",
            "UR"
        }
    ),

    UR整備 = Table.TransformColumns(
        データの展開,
        {
            {
                "UR",
                each
                    if _ = null then
                        null
                    else if Text.Contains(Text.Upper(Text.From(_)), "UR") then
                        "UR"
                    else
                        null,
                type nullable text
            }
        }
    ),

    UR列名変更 = Table.RenameColumns(
        UR整備,
        {{"UR", "取込UR"}},
        MissingField.Ignore
    ),

    区分追加 = Table.AddColumn(
        UR列名変更,
        "区分",
        each if [取込UR] = "UR" then "UR" else "通常",
        type text
    ),

    型の変更 = Table.TransformColumnTypes(
        区分追加,
        {
            {"Name", type text},
            {"担当者", type any},
            {"得意先", type any},
            {"現場", type any},
            {"材料", type any},
            {"数量", type any},
            {"単位", type any},
            {"納品日", type any},
            {"注文状況", type any},
            {"注文日", type any},
            {"現場状況", type any},
            {"チェック", type any},
            {"日時", type any},
            {"取込UR", type text},
            {"区分", type text}
        }
    )
in
    型の変更;

shared #"★変換済み結合" = let
    // 元データ
    班員 = #"★班員データ",

    // 変換リスト側：変換前が空欄の行は除外
    変換元 = 変換リスト,
    変換 = Table.SelectRows(
        変換元,
        each [#"変換前（材料名）"] <> null
            and Text.Trim(Text.From([#"変換前（材料名）"])) <> ""
    ),

    // 班員データ側：材料が空欄ならキーは null、入っていれば正規化
    班員キー付 = Table.AddColumn(
        班員,
        "材料キー",
        each
            if [材料] = null or Text.Trim(Text.From([材料])) = "" then
                null
            else
                fnNormalize(Text.From([材料])),
        type nullable text
    ),

    // 変換リスト側：変換前材料名を正規化してキー化
    変換キー付 = Table.AddColumn(
        変換,
        "変換前キー",
        each fnNormalize(Text.From([#"変換前（材料名）"])),
        type nullable text
    ),

    // 材料キーで左外部結合
    結合 = Table.NestedJoin(
        班員キー付,
        {"材料キー"},
        変換キー付,
        {"変換前キー"},
        "変換結果",
        JoinKind.LeftOuter
    ),

    // 変換結果を展開
    // 変換リスト側のURは、取込側URと区別するため「製品UR」にする
    展開 = Table.ExpandTableColumn(
        結合,
        "変換結果",
        {"変換後（製品名）", "UR", "メーカー"},
        {"変換後製品名", "製品UR", "メーカー"}
    ),

    // 作業用キー列を削除
    材料キー削除 = Table.RemoveColumns(
        展開,
        {"材料キー"}
    ),

    // 見やすい列順に並べ替え
    列の並べ替え = Table.ReorderColumns(
        材料キー削除,
        {
            "Name",
            "担当者",
            "得意先",
            "現場",
            "区分",
            "取込UR",
            "材料",
            "変換後製品名",
            "製品UR",
            "メーカー",
            "数量",
            "単位",
            "納品日",
            "注文状況",
            "注文日",
            "現場状況",
            "チェック",
            "日時"
        },
        MissingField.Ignore
    )
in
    列の並べ替え;

shared #"★変換済みデータ" = let
    元 = #"★変換済み結合",

    入力フォーム表示名追加 = Table.AddColumn(
        元,
        "入力フォーム表示名",
        each
            if [変換後製品名] <> null
                and Text.Trim(Text.From([変換後製品名])) <> ""
            then
                Text.From([変換後製品名])
            else
                Text.From([材料]),
        type text
    ),

    変換状態追加 = Table.AddColumn(
        入力フォーム表示名追加,
        "変換状態",
        each
            if [変換後製品名] <> null
                and Text.Trim(Text.From([変換後製品名])) <> ""
            then
                "変換済"
            else
                "未変換",
        type text
    )
in
    変換状態追加
;

shared #"★未変換一覧" = let
    元 = #"★変換済み結合",
    変換できなかった行 = Table.SelectRows(
        元,
        each [変換後製品名] = null
            or Text.Trim(Text.From([変換後製品名])) = ""
    ),
    必要列のみ = Table.SelectColumns(
        変換できなかった行,
        {"Name", "担当者", "区分", "材料", "数量", "単位", "納品日"},
        MissingField.Ignore
    )
in
    必要列のみ;