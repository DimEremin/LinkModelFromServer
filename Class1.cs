using System;
using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.IO;
using OpenModelFromServer;
using Excel = Microsoft.Office.Interop.Excel;



[TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class OpenModelFromRevitServer : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

        List<string> pathlist = ImportFromExcelToListOfStrings.ReadData("Выберите таблицу Excel, содержащую пути к моделям Revit ");
        List<string> linklist = ImportFromExcelToListOfStrings.ReadData("Выберите таблицу Excel, содержащую пути к связям Revit для вставки в модели");
        List<string> currentpathlist;

        foreach (string linkpath in linklist)
        {
            currentpathlist = pathlist;
            while (currentpathlist.Remove(linkpath)){ } 
            ModelBatchHandler.Run(app, currentpathlist, OperationsWithModel.AddLink, linkpath);
            //RevitLinkInstance t = OperationsWithModel.ReturnObject as RevitLinkInstance;
            
            currentpathlist = null;
        }
        


        return Result.Succeeded;


        //Form1 directoryForm = new Form1();
        //directoryForm.ShowDialog();
        //string formResult = directoryForm.FormResult;
        //MessageBox.Show(formResult);

    }
    

       

    public static class ImportFromExcelToListOfStrings
    {

        public static List<string> ReadData(string title)
        {

            Excel.Application excel = new Excel.Application();
            if (null == excel)
            {
                return null;
            }
            excel.Visible = true;
            string path = GetPath(title);
            Excel.Workbook workbook = excel.Workbooks.Open(path);

            Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;
            if (worksheet == null)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка");
                return null;
            }

            Excel.Range usedRange = worksheet.UsedRange;

            List<string> excelExtactData = new List<string>();

            object obj = new object();

            foreach (Excel.Range cell in usedRange.Cells)
            {
                obj = cell.Value2 as object;
                if (obj != null)
                {
                    string name = obj.ToString();
                    if (name != null)
                    {
                        excelExtactData.Add(name);
                    }
                }
            }
            workbook.Close();
            excel.Quit();
            return excelExtactData;
        }
        public static string GetPath(string title)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel spreadsheet files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*)|*",
                Title = title
            };

            if (dialog.ShowDialog() == true)
            {
                return dialog.FileName;
            }
            return null;
        }
    }
        

    public static class ModelBatchHandler 
    {
        public static bool Run <T>
            (
            Autodesk.Revit.ApplicationServices.Application app, 
            List<string> pathlist, 
            Action<Document, T> operation, 
            T arg
            )
        {
            foreach (string pathstring in pathlist)
            {
                if (!pathstring.Contains(".rvt"))
                {
                    pathlist.Remove(pathstring);
                }
            }

            if (pathlist.Count == 0)
            {
                MessageBox.Show("Спискок моделей пуст");
                return false;
            }

            foreach (string path in pathlist)
            {
                Document newdoc = OpenNotDetached(app, ModelPathUtils.ConvertUserVisiblePathToModelPath(path));

                //ниже можно вставить любой код для пакетной обработки моделей

                operation(newdoc, arg);

                //выше можно вставить любой код для пакетной обработки моделей

                TransactWithCentralOptions transactOptions = new TransactWithCentralOptions();
                SynchronizeWithCentralOptions syncOptions = new SynchronizeWithCentralOptions();
                RelinquishOptions rOptions = new RelinquishOptions(true);
                rOptions.UserWorksets = false;
                syncOptions.SetRelinquishOptions(rOptions);
                newdoc.SynchronizeWithCentral(transactOptions, syncOptions);
                newdoc.Close();
            }

            MessageBox.Show("Готово");

            return true;
        }
        public static Document OpenDetached(Autodesk.Revit.ApplicationServices.Application application, ModelPath modelPath)
        {
            OpenOptions options1 = new OpenOptions();

            options1.DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets;
            Document openedDoc = application.OpenDocumentFile(modelPath, options1);
            
            return openedDoc;
        }

        public static Document OpenNotDetached(Autodesk.Revit.ApplicationServices.Application application, ModelPath modelPath)
        {
            OpenOptions options1 = new OpenOptions();

            options1.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach;
            Document openedDoc = application.OpenDocumentFile(modelPath, options1);

            return openedDoc;
        }
    }
    public static class OperationsWithModel
    {
        public static object ReturnObject;
        public static void AddLink(Document doc, string path)
        {
            ModelPath modelpath = ModelPathUtils.ConvertUserVisiblePathToModelPath(
            @"C:\Revit C#\OpenModelFromServer\Test_EOM.rvt");
            RevitLinkInstance instance = null;
            if (path != null)
            {
                ModelPathUtils.ConvertUserVisiblePathToModelPath(path);


                Transaction transaction = new Transaction(doc);

                transaction.Start("Ошибка");

                //ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
                RevitLinkOptions rlo = new RevitLinkOptions(true);
                var linkType = RevitLinkType.Create(doc, modelpath, rlo);
                instance = RevitLinkInstance.Create(doc, linkType.ElementId);

                transaction.Commit();
            }
            ReturnObject = instance;
        }

        public static void AddWorksets(Document doc, List<string> worksetlist)
        {
            Transaction transaction = new Transaction(doc);

            transaction.Start("Ошибка");

            foreach (string workset in worksetlist)
            {
                if (WorksetTable.IsWorksetNameUnique(doc, workset))
                {
                    Workset.Create(doc, workset);
                }
            }

            transaction.Commit();
        }
    }
    
}



    