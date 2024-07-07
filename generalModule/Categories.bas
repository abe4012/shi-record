Attribute VB_Name = "Categories"
Option Explicit

' 船舶検査記録の各項目名のリストを返す
Function InspectRecCategories()

    Dim Categories As Object
    Set Categories = CreateObject("scripting.Dictionary")
    
    ' 検索カテゴリーのリストを連想配列に設定する
    ' 例) "stat"= 識別用文字列, "状況"= 船舶検査記録上での検索文字列
    ' 注意： "並行" と "併行" は入力ミスが起きやすい
    Categories.Add "stat", "状況"
    Categories.Add "year", "年度"
    Categories.Add "refNum", "№"
    Categories.Add "shipName", "船名"
    Categories.Add "shipType", "船舶種類"
    Categories.Add "owner", "所有者"
    Categories.Add "captainName", "船長名"
    Categories.Add "delegater", "委嘱者"
    Categories.Add "delegateStaff", "担当"
    Categories.Add "inspectType", "検査種類"
    Categories.Add "clause", "約款"
    Categories.Add "receiptDate", "受付日"
    Categories.Add "repNo", "鑑定書番号"
    Categories.Add "repNoCreateDate", "発行日"
    Categories.Add "concurrentInspection", "併行検査"
    Categories.Add "shipyard", "造船所"
    Categories.Add "inspectDate", "立会日"
    Categories.Add "unDocking", "下架・出渠日"
    Categories.Add "prevUndocking", "前回下架・出渠日"
    Categories.Add "prevInspection", "前回検査"
    Categories.Add "prevRepNo", "前回鑑定書"
    Categories.Add "grossT", "総トン数"
    Categories.Add "length", "登録長"
    Categories.Add "breadth", "全幅"
    Categories.Add "depth", "全深／型深"
    Categories.Add "shaftDia", "軸径"
    Categories.Add "propellerNum", "翼数"
    Categories.Add "propellerMaterial", "材質"
    Categories.Add "propellerDia", "ダイヤ"
    Categories.Add "propellerPitch", "ピッチ"
    Categories.Add "accidentDetail", "事故内容"
    Categories.Add "marineAccidentReport", "海難報告書"
    Categories.Add "repairAmount", "修繕額"
    Categories.Add "location", "拠点"
    Categories.Add "kmsStaff", "担当者"
    Categories.Add "inspectDateOther", "施検日(初回以降)"
    Categories.Add "remark", "特記事項"
    
    Set InspectRecCategories = Categories

End Function

