using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
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
                        ResultBuffer xdata = entity.GetXDataForApplication("AcDb");
                        List<string> dataList = new List<string>();

                        if (xdata != null)
                        {
                            foreach (TypedValue value in xdata)
                            {
                                dataList.Add($"Type: {value.TypeCode}, Value: {value.Value}");
                            }
                        }
                        else
                        {
                            dataList.Add("このオブジェクトには拡張データがありません。");
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
                ed.WriteMessage($"\nエラーが発生しました: {ex.Message}");
            }
        }
    }
}
