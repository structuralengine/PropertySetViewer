using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

[assembly: CommandClass(typeof(PropertySetViewer.PropertySetViewer))]

namespace PropertySetViewer
{
    public class PropertySetViewer
    {
        [CommandMethod("ViewPropertySetGUI")]
        public void ViewPropertySetGUI()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // オブジェクトの選択
                PromptEntityOptions entityOptions = new PromptEntityOptions("\n拡張データを確認するオブジェクトを選択してください: ");
                PromptEntityResult entityResult = ed.GetEntity(entityOptions);

                if (entityResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nオブジェクトの選択がキャンセルされました。");
                    return;
                }

                // データベーストランザクションの開始
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = tr.GetObject(entityResult.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                    if (entity != null)
                    {
                        List<string> dataList = new List<string>();
                        bool dataFound = false;

                        // PLACEHOLDER: Variable declarations for dataList and dataFound remain unchanged

                        // 拡張辞書を確認
                        ObjectId extDictId = entity.ExtensionDictionary;
                        if (!extDictId.IsNull)
                        {
                            using (DBDictionary extDict = tr.GetObject(extDictId, OpenMode.ForRead) as DBDictionary)
                            {
                                if (extDict != null)
                                {
                                    // 特定のプロパティセット名を定義
                                    string[] propertySetNames = new string[] { "施工情報(一覧表)", "施工情報(個別)" };

                                    foreach (DBDictionaryEntry entry in extDict)
                                    {
                                        // プロパティセット名が一致する場合のみ処理
                                        if (propertySetNames.Contains(entry.Key))
                                        {
                                            using (Autodesk.AutoCAD.DatabaseServices.DBObject obj = tr.GetObject(entry.Value, OpenMode.ForRead))
                                            {
                                                dataFound = true;
                                                dataList.Add($"プロパティセット: {entry.Key}");
                                                if (obj is Xrecord xrec)
                                                {
                                                    foreach (TypedValue value in xrec.Data)
                                                    {
                                                        string propertyName = GetPropertyName(value.TypeCode);
                                                        dataList.Add($"  {propertyName}: {value.Value}");
                                                    }
                                                }
                                                dataList.Add(""); // 空行を追加
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // XData を確認
                        string[] appNames = new string[] { "CIVIL", "CIVILDATA", "PROPERTYSETS" };
                        foreach (string appName in appNames)
                        {
                            ResultBuffer xdata = entity.GetXDataForApplication(appName);
                            if (xdata != null)
                            {
                                dataFound = true;
                                dataList.Add($"XData ({appName}):");
                                foreach (TypedValue value in xdata)
                                {
                                    string propertyName = GetPropertyName(value.TypeCode);
                                    dataList.Add($"  {propertyName}: {value.Value}");
                                }
                                dataList.Add("");
                            }
                        }

                        if (!dataFound)
                        {
                            dataList.Add("このオブジェクトには施工情報の拡張データ（プロパティセット、拡張辞書、XData）が見つかりませんでした。");
                        }

                        // プロパティ名を取得するヘルパーメソッド
                        string GetPropertyName(int typeCode)
                        {
                            switch (typeCode)
                            {
                                case 1: return "改良住番号";
                                case 2: return "設計統芯位置";
                                case 3: return "施工実績";
                                case 4: return "コラム番号";
                                case 5: return "施工時間";
                                case 6: return "区間開始深さ";
                                case 7: return "区間終了深さ";
                                case 8: return "施工実績x";
                                case 9: return "施工実績y";
                                case 10: return "設計統芯位置x";
                                case 11: return "設計統芯位置y";
                                case 1001: return "施工データ";
                                default: return $"プロパティ{typeCode}";
                            }
                        }

                        // GUIを表示
                        PropertySetForm form = new PropertySetForm(dataList);
                        Application.ShowModalDialog(form);

                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nエラーが発生しました: {ex.Message}\n拡張データの取得に失敗しました。");
                }
            }
        }
    }
}
