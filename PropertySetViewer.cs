using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using System.Collections.Generic;
using System.Xml.Linq;

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
                    Entity entity = tr.GetObject(entityResult.ObjectId, OpenMode.ForRead) as Entity;
                    if (entity != null)
                    {
                        List<string> dataList = new List<string>();
                        bool dataFound = false;

                        // Civil3D のプロパティセットを確認
                        if (PropertySetManager.HasPropertySets(entity))
                        {
                            var propertySets = PropertySetManager.GetAllPropertySets(entity);
                            foreach (var propertySet in propertySets)
                            {
                                // 特定のプロパティセットを探す
                                if (propertySet.PropertySetDefinitionName == "施工情報(一覧表)" ||
                                    propertySet.PropertySetDefinitionName == "施工情報(個別)")
                                {
                                    dataFound = true;
                                    dataList.Add($"プロパティセット: {propertySet.PropertySetDefinitionName}");
                                    foreach (var definition in propertySet.Definitions)
                                    {
                                        var value = propertySet.GetPropertyValue(definition);
                                        dataList.Add($"  {definition.Name}: {value}");
                                    }
                                    dataList.Add(""); // 空行を追加して見やすくする
                                }
                            }
                        }

                        // 拡張辞書を確認
                        ObjectId extDictId = entity.ExtensionDictionary;
                        if (!extDictId.IsNull)
                        {
                            using (DBDictionary extDict = tr.GetObject(extDictId, OpenMode.ForRead) as DBDictionary)
                            {
                                if (extDict != null)
                                {
                                    foreach (DBDictionaryEntry entry in extDict)
                                    {
                                        using (DBObject obj = tr.GetObject(entry.Value, OpenMode.ForRead))
                                        {
                                            dataFound = true;
                                            dataList.Add($"拡張辞書エントリ: {entry.Key}");
                                            if (obj is Xrecord xrec)
                                            {
                                                foreach (TypedValue value in xrec.Data)
                                                {
                                                    dataList.Add($"  Type: {value.TypeCode}, Value: {value.Value}");
                                                }
                                            }
                                            dataList.Add("");
                                        }
                                    }
                                }
                            }
                        }

                        // XData を確認
                        ResultBuffer xdata = entity.GetXDataForApplication("CIVIL");
                        if (xdata != null)
                        {
                            dataFound = true;
                            dataList.Add("XData:");
                            foreach (TypedValue value in xdata)
                            {
                                dataList.Add($"  Type: {value.TypeCode}, Value: {value.Value}");
                            }
                            dataList.Add("");
                        }

                        if (!dataFound)
                        {
                            dataList.Add("このオブジェクトには施工情報の拡張データ（プロパティセット、拡張辞書、XData）が見つかりませんでした。");
                        }

                        // GUIを表示
                        PropertySetForm form = new PropertySetForm(dataList);
                        Application.ShowModalDialog(form);
                    }

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
