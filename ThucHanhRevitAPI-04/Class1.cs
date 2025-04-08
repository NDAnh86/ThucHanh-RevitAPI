using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using System.Xml;

namespace ThucHanhRevitAPI_04
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class TH04 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Document doc = uidoc.Document;
            string modelPath = doc.PathName;
            string modelTitle = doc.Title;
            StreamWriter sw = new StreamWriter(modelPath.Replace(doc.Title + ".rvt", "Result.txt"));
            sw.WriteLine("BaseConstraint\tVolume\tArea\tLength");
            foreach (ElementId elemId in ids)
            {
                Element elem = doc.GetElement(elemId);
                Wall wall = elem as Wall;
                if (wall != null)
                {
                    string condition = findCondition(wall, doc);
                    double volume = findVolume(wall);
                    double area = findArea(wall);
                    double length = findLength(wall);
                    sw.WriteLine(condition + "\t" + volume + "\t" + area + "\t" + length);
                }
            }
            sw.Close();
            MessageBox.Show("Đã lưu file dữ liệu của tường thành công");
            


            Transaction trans = new Transaction(doc);
            trans.Start("instance binding");

            
            FileStream fileStream = File.Create("D:\\ShareParameter.txt");
            fileStream.Close();

            
            doc.Application.SharedParametersFilename = "D:\\ShareParameter.txt";
            DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
            bool bindResult = SetNewParameterToTypeWall(uidoc.Application, defFile);

            
            if (bindResult == true)
            {
                MessageBox.Show("Successfully add the new parameter \"Note\" by instance binding!");
            }

            trans.Commit();


            Autodesk.Revit.UI.Selection.Selection selElement = uidoc.Selection;
            Autodesk.Revit.DB.Document doc2 = uidoc.Document;
            Transaction trans2 = new Transaction(doc);
            trans2.Start("set values for parameters");
            int k = 1;
            foreach (ElementId eleId in selElement.GetElementIds())
            {
                Element ele = doc2.GetElement(eleId);
                this.setMark(k, eleId, doc2);
                k++;
            }
            trans2.Commit();
            MessageBox.Show("Set dữ liệu thành công");
            return Result.Succeeded;


        }

        public bool SetNewParameterToTypeWall(UIApplication app, DefinitionFile myDefinitionFile)
        {
            // Create a new group in the shared parameters file
            DefinitionGroups myGroups = myDefinitionFile.Groups;
            DefinitionGroup myGroup = myGroups.Create("Revit API Course");
            // Create a type definition
            ExternalDefinitionCreationOptions option = new ExternalDefinitionCreationOptions("Note", SpecTypeId.String.Text);
            Definition myDefinition_CompanyName = myGroup.Definitions.Create(option);
            // Create a category set and insert category of wall to it
            CategorySet myCategories = app.Application.Create.NewCategorySet();
            // Use BuiltInCategory to get category of wall
            Category myCategory = Category.GetCategory(app.ActiveUIDocument.Document, BuiltInCategory.OST_Walls);
            myCategories.Insert(myCategory);//add wall into the group. Of course, you can add multiple categories

            InstanceBinding instanceBinding = app.Application.Create.NewInstanceBinding(myCategories);
            // Get the BingdingMap of current document.
            BindingMap bindingMap = app.ActiveUIDocument.Document.ParameterBindings;
            //Bind the definitions to the document
            bool typeBindOK = bindingMap.Insert(myDefinition_CompanyName, instanceBinding, GroupTypeId.Text);
            return typeBindOK;
        }


        public void setMark(int k, ElementId eleId, Document doc2)
        {
            string message = "Đây là đối tượng thứ " + k + " đã được lựa chọn";
            Element ele = doc2.GetElement(eleId);
            foreach (Parameter para in ele.Parameters)
            {
                if (para.Definition.Name == "Note")
                {
                    try
                    {
                        para.Set(message);
                    }
                    catch (Exception ex) { }
                }
            }
        }


        private double findVolume(Element elem)
        {
            double volume = 0.0;
            volume = elem.LookupParameter("Volume").AsDouble();
            return volume;
        }

        private string findCondition(Element elem, Document doc)
        {
            string name = null;
            name = doc.GetElement(elem.LookupParameter("Base Constraint").AsElementId()).Name;
            return name;
        }

        private double findArea(Element elem)
        {
            double area = 0.0;
            area = elem.LookupParameter("Area").AsDouble();
            return area;
        }

        private double findLength(Element elem)
        {
            double length = 0.0;
            length = elem.LookupParameter("Length").AsDouble();
            return length;
        }






    }




}
