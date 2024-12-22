using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using Exception = System.Exception;

[assembly: CommandClass(typeof(PropertySetViewer.PropertySetViewer))]

namespace PropertySetViewer
{
    public class PropertySetViewer
    {
        private string DecodeTypedValue(TypedValue value)
        {
            try
            {
                switch (value.TypeCode)
                {
                    case 1:  // Text
                    case 2:  // Name
                    case 3:  // Block name
                        if (value.Value == null) return "null";
                        return System.Text.Encoding.UTF8.GetString(
                            System.Text.Encoding.GetEncoding("Shift-JIS")
                            .GetBytes(value.Value.ToString()));

                    case 40:  // Real
                    case 50:  // Angle
                        return value.Value is double d ? d.ToString("F6") : "0.000000";

                    case 70:  // Integer
                    case 90:  // 32-bit integer
                    case 1071: // Extended integer
                        return value.Value?.ToString() ?? "0";

                    case -3:  // Extended data
                        return ProcessExtendedData(value.Value);

                    default:
                        return value.Value?.ToString() ?? "null";
                }
            }
            catch (System.Exception ex)
            {
                return $"デコードエラー: {ex.Message}";
            }
        }

        private string ProcessExtendedData(object data)
        {
            if (data == null) return "null";

            try
            {
                // Handle ACAD_STEPID and other special cases
                if (data is ResultBuffer rb)
                {
                    var values = new List<string>();
                    foreach (TypedValue tv in rb)
                    {
                        // Special handling for binary data
                        if (tv.TypeCode == 1004) // Binary chunk
                        {
                            if (tv.Value is byte[] bytes)
                            {
                                try
                                {
                                    // Try UTF-8 first
                                    string text = System.Text.Encoding.UTF8.GetString(bytes);
                                    if (!string.IsNullOrWhiteSpace(text) && text.All(c => !char.IsControl(c) || c == '\n' || c == '\r'))
                                    {
                                        values.Add(text);
                                        continue;
                                    }
                                }
                                catch { }

                                try
                                {
                                    // Try Shift-JIS if UTF-8 fails
                                    string text = System.Text.Encoding.GetEncoding("Shift-JIS").GetString(bytes);
                                    if (!string.IsNullOrWhiteSpace(text) && text.All(c => !char.IsControl(c) || c == '\n' || c == '\r'))
                                    {
                                        values.Add(text);
                                        continue;
                                    }
                                }
                                catch { }

                                // If text decoding fails, show improved hex representation
                                values.Add($"バイナリデータ: {BitConverter.ToString(bytes).Replace("-", " ")}");
                            }
                            continue;
                        }

                        values.Add(DecodeTypedValue(tv));
                    }
                    return string.Join(", ", values);
                }

                // Handle other types of extended data
                if (data is TypedValue tv2)
                {
                    return DecodeTypedValue(tv2);
                }

                return data.ToString();
            }
            catch (Exception ex)
            {
                return $"拡張データ処理エラー: {ex.Message}";
            }
        }

        private void ProcessXrecordData(Xrecord xrec, List<string> dataList)
        {
            if (xrec?.Data == null) return;

            foreach (TypedValue value in xrec.Data)
            {
                string propertyName = GetPropertyName(value.TypeCode);
                string propertyValue = DecodeTypedValue(value);
                dataList.Add($"  {propertyName}: {propertyValue}");
            }
        }

        private void ProcessExtensionDictionary(Entity entity, Transaction tr, List<string> dataList, ref bool dataFound)
        {
            ObjectId extDictId = entity.ExtensionDictionary;
            if (extDictId.IsNull)
            {
                return;
            }

            try
            {
                using (DBDictionary extDict = tr.GetObject(extDictId, OpenMode.ForRead) as DBDictionary)
                {
                    if (extDict != null)
                    {
                        string[] propertySetNames = new string[] { "施工情報(一覧表)", "施工情報(個別)" };

                        foreach (DBDictionaryEntry entry in extDict)
                        {
                            if (propertySetNames.Contains(entry.Key))
                            {
                                using (DBObject obj = tr.GetObject(entry.Value, OpenMode.ForRead))
                                {
                                    dataFound = true;
                                    dataList.Add($"プロパティセット: {entry.Key}");
                                    if (obj is Xrecord xrec)
                                    {
                                        ProcessXrecordData(xrec, dataList);
                                    }
                                    dataList.Add("");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dataList.Add($"拡張ディクショナリの処理中にエラーが発生しました: {ex.Message}");
            }
        }

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

                        // 拡張辞書を確認
                        ProcessExtensionDictionary(entity, tr, dataList, ref dataFound);

                        // XData を確認
                        string[] appNames = new string[] { "CIVIL", "CIVILDATA", "PROPERTYSETS", "CIVIL3D", "C3D", "AEC" };
                        foreach (string appName in appNames)
                        {
                            ResultBuffer xdata = entity.GetXDataForApplication(appName);
                            if (xdata != null)
                            {
                                dataFound = true;
                                dataList.Add($"XData ({appName}):");
                                foreach (TypedValue value in xdata)
                                {
                                    var rb = new ResultBuffer(new TypedValue[] { value });
                                    ProcessXrecordData(new Xrecord { Data = rb }, dataList);
                                }
                                dataList.Add("");
                            }
                        }

                        if (!dataFound)
                        {
                            dataList.Add("このオブジェクトには施工情報の拡張データ（プロパティセット、拡張辞書、XData）が見つかりませんでした。");
                            dataList.Add("\nデバッグ情報:");
                            dataList.Add($"オブジェクトタイプ: {entity.GetType().Name}");
                            dataList.Add($"拡張辞書ID: {entity.ExtensionDictionary}");
                            var xdataApps = appNames.Where(appName => entity.GetXDataForApplication(appName) != null);
                            dataList.Add($"XData アプリケーション名: {string.Join(", ", xdataApps.Any() ? xdataApps : new string[] { "なし" })}");
                        }

                        // GUIを表示
                        PropertySetForm form = new PropertySetForm(dataList);
                        Application.ShowModalDialog(form);

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nエラーが発生しました: {ex.Message}");
                ed.WriteMessage($"\nスタックトレース: {ex.StackTrace}");
                ed.WriteMessage("\n拡張データの取得に失敗しました。");
            }
        }

        private string GetPropertyName(int typeCode)
        {
            switch (typeCode)
            {
                case -3: return "拡張データ";
                case 1: return "テキストデータ";
                case 2: return "名前";
                case 3: return "ブロック名";
                case 40: return "実数値";
                case 70: return "整数値";
                case 90: return "32ビット整数";
                case 1071: return "拡張整数";
                // 施工情報の特殊プロパティ
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
    }
}
