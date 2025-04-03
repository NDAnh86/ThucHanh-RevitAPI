using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows.Forms;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;

namespace Ex_TaoListElement_Filter
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            ICollection<ElementId> selectIds = uidoc.Selection.GetElementIds();

            TaskDialog.Show("Revit", "Số lượng đối tượng được chọn: " + selectIds.Count.ToString());
            // đếm số lượng cột đã chọn
            ICollection<ElementId> selectedWallIds = new List<ElementId>();
            foreach (ElementId id in selectIds)
            {
                Element elements1 = uidoc.Document.GetElement(id);
                if (elements1 is Wall)
                {
                    selectedWallIds.Add(id);
                }
            }
            if(0 != selectedWallIds.Count)
            {
                TaskDialog.Show("Revit", selectedWallIds.Count.ToString() + " Đối tượng tường đã được chọn!");
            }
            else
            {
                TaskDialog.Show("Revit", selectedWallIds.Count.ToString() + "Không có đối tượng tường nào được chọn!");
            }

            // đếm số lượng dầm đã chọn
            ICollection<ElementId> selectedFloorIds = new List<ElementId>();
            foreach (ElementId id in selectIds)
            {
                Element elements2 = uidoc.Document.GetElement(id);
                if (elements2 is Floor)
                {
                    selectedFloorIds.Add(id);
                }
            }
            if (0 != selectedFloorIds.Count)
            {
                TaskDialog.Show("Revit", selectedFloorIds.Count.ToString() + " Đối tượng sàn đã được chọn!");
            }
            else
            {
                TaskDialog.Show("Revit", selectedFloorIds.Count.ToString() + "Không có đối tượng sàn nào được chọn!");
            }

            int columnCount = selectIds.Select(id => uidoc.Document.GetElement(id))
                                   .Count(el => el is FamilyInstance fi &&
                                                fi.StructuralType == StructuralType.Column);

            int beamCount = selectIds.Select(id => uidoc.Document.GetElement(id))
                                   .Count(el => el is FamilyInstance fi &&
                                                fi.StructuralType == StructuralType.Beam);




            ///VD dùng Filter
            //cột

            Document doc = uidoc.Document;

            FilteredElementCollector collector = new FilteredElementCollector(doc);

            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);

            IList<Element> columnList = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
            StringBuilder output = new StringBuilder("Tất cả các Structural Column trong BIM model :\r\n");

            foreach (Element elem in columnList)
            {
                FamilyInstance familyInstance = elem as FamilyInstance;
                FamilySymbol familySymbol = familyInstance.Symbol;
                Family family = familySymbol.Family;
                string elemName = "Category:" + elem.Category.Name + ", Family:" + family.Name + ", Name:" + elem.Name + "\r\n";
                output.Append(elemName);
            }
            MessageBox.Show(output.ToString());

            //dầm
            Document doc1 = uidoc.Document;

            FilteredElementCollector collector1 = new FilteredElementCollector(doc1);

            ElementCategoryFilter filter1 = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);

            IList<Element> framingList = collector1.WherePasses(filter1).WhereElementIsNotElementType().ToElements();
            StringBuilder output1 = new StringBuilder("Tất cả các Structural Framing trong BIM model :\r\n");
            foreach (Element elem1 in framingList)
            {
                FamilyInstance familyInstance1 = elem1 as FamilyInstance;
                FamilySymbol familySymbol1 = familyInstance1.Symbol;
                Family family1 = familySymbol1.Family;
                string elemName1 = "Category:" + elem1.Category.Name + ", Family:" + family1.Name + ", Name:" + elem1.Name + "\r\n";
                output1.Append(elemName1);
            }
            MessageBox.Show(output1.ToString());


            //cửa sổ
            Document doc3 = uidoc.Document;

            FilteredElementCollector collector3 = new FilteredElementCollector(doc3);

            ElementCategoryFilter filter3 = new ElementCategoryFilter(BuiltInCategory.OST_Windows);



            IList<Element> windowList = collector3.WherePasses(filter3).WhereElementIsNotElementType().ToElements();
            StringBuilder output3 = new StringBuilder("Tất cả các Window trong BIM model :\r\n");
            foreach (Element elem3 in windowList)
            {
                FamilyInstance familyInstance3 = elem3 as FamilyInstance;
                FamilySymbol familySymbol3 = familyInstance3.Symbol;
                Family family3 = familySymbol3.Family;
                string elemName3 = "Category:" + elem3.Category.Name + ", Family:" + family3.Name + ", Name:" + elem3.Name + "\r\n";
                output3.Append(elemName3);
            }
            MessageBox.Show(output3.ToString());

            int windowCount = collector3.Count();


            Document doc2 = commandData.Application.ActiveUIDocument.Document;
            FilteredElementCollector collector2 = new FilteredElementCollector(doc2);

            ElementCategoryFilter filter2 = new ElementCategoryFilter(BuiltInCategory.OST_Materials);
            IList<Element> matList = collector2.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            StringBuilder result = new StringBuilder();
            foreach (Element ele in matList)
            {
                result.AppendLine(ele.Category.Name + "\t" + ele.Name + "\t" + ele.Id);
            }

            MessageBox.Show("Tất cả các materials trong project là:\r\n" + result.ToString());


            MessageBox.Show("Số lượng tường: " + selectedWallIds.Count.ToString() + "\n"
                            + "Số lượng cột: " + columnCount.ToString() + "\n"
                            + "Số lượng dầm: " + beamCount.ToString() + "\n"
                            + "Số lượng sàn: " + selectedFloorIds.Count.ToString() + "\n"
                            + "Số lượng cửa sổ: " + windowCount.ToString());







            return Result.Succeeded;
        }
       
    }
}
