using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Aec.DatabaseServices;
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

        private void ProcessDictionaryObject(DBObject dictObj, List<string> dataList)
        {
            if (dictObj == null) return;

            try
            {
                // プロパティセットの内容を解析
                if (dictObj is PropertySet propSet)
                {
                    foreach (PropertySetProperty prop in propSet)
                    {
                        string value = prop.PropertyValue?.ToString() ?? "null";
                        try
                        {
                            // バイナリデータの場合は特別な処理を試みる
                            if (prop.PropertyValue is byte[] bytes)
                            {
                                value = ProcessBinaryData(bytes);
                            }
                        }
                        catch (Exception ex)
                        {
                            value = $"値の変換エラー: {ex.Message}";
                        }
                        dataList.Add($"  {prop.PropertyName}: {value}");
                    }
                }
                else if (dictObj is Xrecord xrec)
                {
                    ProcessXrecordData(xrec, dataList);
                }
                else
                {
                    // その他の辞書オブジェクトの情報を表示
                    dataList.Add($"  タイプ: {dictObj.GetType().Name}");
                    dataList.Add($"  オブジェクトID: {dictObj.ObjectId}");

                    // オブジェクトのプロパティを取得して表示
                    var props = dictObj.GetType().GetProperties()
                        .Where(p => p.CanRead && !p.Name.Equals("ObjectId") && !p.Name.Equals("Handle"));

                    foreach (var prop in props)
                    {
                        try
                        {
                            var value = prop.GetValue(dictObj);
                            if (value != null)
                            {
                                dataList.Add($"    {prop.Name}: {value}");
                            }
                        }
                        catch { } // プロパティの読み取りに失敗した場合はスキップ
                    }
                }
            }
            catch (Exception ex)
            {
                dataList.Add($"  辞書オブジェクト処理エラー: {ex.Message}");
                dataList.Add($"  スタックトレース: {ex.StackTrace}");
            }
        }

        private string ProcessBinaryData(byte[] bytes)
        {
            try
            {
                // UTF-8でのデコードを試みる
                string utf8Text = System.Text.Encoding.UTF8.GetString(bytes);
                if (!string.IsNullOrWhiteSpace(utf8Text) && utf8Text.All(c => !char.IsControl(c) || c == '\n' || c == '\r'))
                {
                    return utf8Text;
                }

                // Shift-JISでのデコードを試みる
                string sjisText = System.Text.Encoding.GetEncoding("Shift-JIS").GetString(bytes);
                if (!string.IsNullOrWhiteSpace(sjisText) && sjisText.All(c => !char.IsControl(c) || c == '\n' || c == '\r'))
                {
                    return sjisText;
                }

                // バイナリデータとして表示
                return $"バイナリデータ: {BitConverter.ToString(bytes).Replace("-", " ")}";
            }
            catch (Exception ex)
            {
                return $"バイナリデータ処理エラー: {ex.Message}";
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
                        // Civil3D、AEC関連のプロパティセット名を拡張
                        string[] propertySetNames = new string[] {
                            "施工情報(一覧表)",
                            "施工情報(個別)",
                            "CIVIL",
                            "CIVILDATA",
                            "PROPERTYSETS",
                            "CIVIL3D",
                            "C3D",
                            "AEC"
                        };

                        // すべての辞書エントリを列挙（デバッグ用）
                        var allEntries = new List<string>();
                        foreach (DBDictionaryEntry entry in extDict)
                        {
                            allEntries.Add(entry.Key);
                        }
                        dataList.Add($"利用可能な辞書エントリ: {string.Join(", ", allEntries)}");

                        // プロパティセットの処理
                        foreach (DBDictionaryEntry entry in extDict)
                        {
                            bool isRelevantEntry = propertySetNames.Any(name =>
                                entry.Key.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);

                            if (isRelevantEntry)
                            {
                                using (DBObject obj = tr.GetObject(entry.Value, OpenMode.ForRead))
                                {
                                    dataFound = true;
                                    dataList.Add($"プロパティセット: {entry.Key}");
                                    if (obj is Xrecord xrec)
                                    {
                                        ProcessXrecordData(xrec, dataList);
                                    }
                                    else
                                    {
                                        ProcessDictionaryObject(obj, dataList);
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
                dataList.Add($"スタックトレース: {ex.StackTrace}");
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

                            // 拡張辞書エントリの詳細情報を追加
                            if (entity.ExtensionDictionary != ObjectId.Null)
                            {
                                try
                                {
                                    using (DBDictionary extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary)
                                    {
                                        if (extDict != null)
                                        {
                                            var entries = new List<string>();
                                            foreach (DBDictionaryEntry entry in extDict)
                                            {
                                                using (DBObject obj = tr.GetObject(entry.Value, OpenMode.ForRead))
                                                {
                                                    entries.Add($"{entry.Key}（タイプ: {obj.GetType().Name}）");
                                                }
                                            }
                                            dataList.Add($"拡張辞書エントリ: {string.Join(", ", entries)}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    dataList.Add($"拡張辞書エントリの取得エラー: {ex.Message}");
                                }
                            }
                            else
                            {
                                dataList.Add("拡張辞書エントリ: なし");
                            }

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

        private string SerializeObjectData(DBObject obj, Transaction tr, int indent = 0)
        {
            var builder = new System.Text.StringBuilder();
            string padding = new string(' ', indent * 2);

            builder.AppendLine($"{padding}{{");
            builder.AppendLine($"{padding}  \"Type\": \"{obj.GetType().Name}\",");
            builder.AppendLine($"{padding}  \"ObjectId\": \"{obj.ObjectId}\",");

            if (obj is DBDictionary dict)
            {
                builder.AppendLine($"{padding}  \"Entries\": {{");
                foreach (DBDictionaryEntry entry in dict)
                {
                    builder.AppendLine($"{padding}    \"{entry.Key}\": {{");
                    using (DBObject entryObj = tr.GetObject(entry.Value, OpenMode.ForRead))
                    {
                        builder.Append(SerializeObjectData(entryObj, tr, indent + 3));
                    }
                    builder.AppendLine($"{padding}    }}}},");
                }
                builder.AppendLine($"{padding}  }}");
            }
            else if (obj is Xrecord xrec)
            {
                builder.AppendLine($"{padding}  \"Data\": [");
                foreach (TypedValue value in xrec.Data)
                {
                    builder.AppendLine($"{padding}    {{");
                    builder.AppendLine($"{padding}      \"TypeCode\": {value.TypeCode},");
                    builder.AppendLine($"{padding}      \"TypeName\": \"{GetPropertyName(value.TypeCode)}\",");
                    builder.AppendLine($"{padding}      \"Value\": \"{DecodeTypedValue(value)}\"");
                    builder.AppendLine($"{padding}    }},");
                }
                builder.AppendLine($"{padding}  ]");
            }

            builder.AppendLine($"{padding}}}");
            return builder.ToString();
        }

        [CommandMethod("ExportPropertySetData")]
        public void ExportPropertySetData()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptEntityOptions entityOptions = new PromptEntityOptions("\n拡張データをエクスポートするオブジェクトを選択してください: ");
                PromptEntityResult entityResult = ed.GetEntity(entityOptions);

                if (entityResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nオブジェクトの選択がキャンセルされました。");
                    return;
                }

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Entity entity = tr.GetObject(entityResult.ObjectId, OpenMode.ForRead) as Entity;
                    if (entity != null)
                    {
                        var builder = new System.Text.StringBuilder();
                        builder.AppendLine("{");
                        builder.AppendLine($"  \"EntityType\": \"{entity.GetType().Name}\",");
                        builder.AppendLine($"  \"ExtensionDictionary\": {{");

                        if (!entity.ExtensionDictionary.IsNull)
                        {
                            using (DBDictionary extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary)
                            {
                                builder.Append(SerializeObjectData(extDict, tr, 2));
                            }
                        }

                        builder.AppendLine("  },");
                        builder.AppendLine("  \"XData\": {");

                        string[] appNames = { "CIVIL", "CIVILDATA", "PROPERTYSETS", "CIVIL3D", "C3D", "AEC" };
                        foreach (string appName in appNames)
                        {
                            ResultBuffer xdata = entity.GetXDataForApplication(appName);
                            if (xdata != null)
                            {
                                builder.AppendLine($"    \"{appName}\": {{");
                                builder.AppendLine("      \"Data\": [");
                                foreach (TypedValue value in xdata)
                                {
                                    builder.AppendLine("        {");
                                    builder.AppendLine($"          \"TypeCode\": {value.TypeCode},");
                                    builder.AppendLine($"          \"TypeName\": \"{GetPropertyName(value.TypeCode)}\",");
                                    builder.AppendLine($"          \"Value\": \"{DecodeTypedValue(value)}\"");
                                    builder.AppendLine("        },");
                                }
                                builder.AppendLine("      ]");
                                builder.AppendLine("    },");
                            }
                        }

                        builder.AppendLine("  }");
                        builder.AppendLine("}");

                        string fileName = $"PropertySetData_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                        string filePath = System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                            fileName
                        );
                        System.IO.File.WriteAllText(filePath, builder.ToString());

                        ed.WriteMessage($"\nデータを保存しました: {filePath}");
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nエラーが発生しました: {ex.Message}");
                ed.WriteMessage($"\nスタックトレース: {ex.StackTrace}");
            }
        }
    }
}
