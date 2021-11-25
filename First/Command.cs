#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace First
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CreatePipeButton : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            App.m_App = commandData.Application;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document document = uidoc.Document;

            // ������ Line ��������
            List<Element> elementlist = GetPipeLine(uidoc, document);

            // Pipe ���� �� Pipe Fitting
            CreatePipe(uidoc, document, elementlist);
            return Result.Succeeded;
        }




        public class PlanarFacesSelectionFilter : ISelectionFilter
        {
            Document doc = null;
            public PlanarFacesSelectionFilter(Document document)
            {
                doc = document;
            }

            public bool AllowElement(Element element)
            {
                return true;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                if (doc.GetElement(refer).GetGeometryObjectFromReference(refer) is PlanarFace)
                {
                    return true;

                }
                return false;
            }

        }


        public List<Element> GetPipeLine(UIDocument uidoc, Document document)
        {
            ISelectionFilter selFilter = new PlanarFacesSelectionFilter(document);
            IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "Select multiple planar faces");
            List<Element> elementlist = new List<Element>();

            foreach (Reference referen in references)
            {
                Element element = uidoc.Document.GetElement(referen);
                elementlist.Add(element);
            }

            return elementlist;
        }


        public void Createpipe(UIDocument uidoc, Document document, List<Element> elementlist)
        {

            // (1) pipe �����ϱ�

            // 1-1) pipe ������ �ʿ��� �Ӽ� ���Ƿ� ��������
            PipeType pipeType = new FilteredElementCollector(document).OfClass(typeof(PipeType)).FirstElement() as PipeType;
            Level level = new FilteredElementCollector(document).OfClass(typeof(Level)).First() as Level;
            FilteredElementCollector sysCollector = new FilteredElementCollector(document);
            sysCollector.OfClass(typeof(PipingSystemType));
            ElementId pipeSysTypeId = sysCollector.FirstElementId();
            bool flag = true;

            // 1-2) ��������
            Pipe newpipe = null;
            XYZ start = null;
            XYZ end = null;

            // 1-3) geometryElements �����ϱ�
            List<GeometryElement> geometryElements = new List<GeometryElement>();

            foreach (Element element in elementlist)
            {
                GeometryElement geometry = element.get_Geometry(new Options());
                geometryElements.Add(geometry);
            }

            // 1-4) Ʈ������ ����
            using (Transaction trans = new Transaction(document))
            {
                try
                {
                    trans.Start("Create pipe");

                    // 1-5) ������ ���� �� ����Ʈ�� ���
                    List<Line>lines = new List<Line>();
                    List<Pipe>pipes = new List<Pipe>();
                    List<Element>elePipe = new List<Element>();
                    foreach (GeometryElement geometryele in geometryElements)
                    {
                        foreach (GeometryObject obj in geometryele)
                        {
                            Line line = obj as Line;

                            lines.Add(line);


                            start = line.GetEndPoint(0);
                            end = line.GetEndPoint(1);
                            if (pipeType != null)
                            {

                                newpipe = Pipe.Create(document, pipeSysTypeId, pipeType.Id, level.Id, start, end);
                                pipes.Add(newpipe);

                                Element element = document.GetElement(newpipe.Id as ElementId);
                                elePipe.Add(element);

                                // 1-6) fitting �� elbow ���⿡ �°� pipe ���� �����ϱ�
                                ElementId elementId = newpipe.Id as ElementId;
                                Parameter parameter = element.LookupParameter("Diameter");
                                parameter.Set(10 * 0.007333);

                                // 1-7) Pipe �� ������ elbow Type �����ϱ�
                                // Revit api �� pipe fitting �� ��� ������ elbowType�� �⺻���� none�̱� ������ ��! Routing Preferences���� �������־���Ѵ�.
                                ElementType elbowType = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_PipeFitting).OfClass(typeof(ElementType)).Cast<ElementType>().Where(x => x.FamilyName.Contains("M_Elbow")).FirstOrDefault();
                                RoutingPreferenceManager rpm = newpipe.PipeType.RoutingPreferenceManager;
                                rpm.AddRule(RoutingPreferenceRuleGroupType.Elbows, new RoutingPreferenceRule(elbowType.Id, "Set Elbow fitting Type"));
                                int routingPerenceGroupCnt = rpm.GetNumberOfRules(RoutingPreferenceRuleGroupType.Elbows);
                                if (routingPerenceGroupCnt>1)
                                {
                                    for (int k = 0; k<routingPerenceGroupCnt - 1; k++)
                                    {
                                        rpm.RemoveRule(RoutingPreferenceRuleGroupType.Elbows, 0);
                                    }
                                }
                            }
                        }
                    }

                    // (2) Pipe �����Ű�� 

                    // 2-1) transaction �ȿ��� ���� ���������Ƿ� �߰��� ���ΰ�ħ�Ͽ� elbow ���� ������ �� ����ǵ��� �ϱ� 
                    document.Regenerate();
                    // 2-2) �����ų 2���� ������ connector ������������, connectormanager�� connectors �̾Ƴ���

                    for (int i = 0; i< pipes.Count() - 1; i++)
                    {

                        ConnectorManager pipe_connectorManager1 = pipes[i].ConnectorManager;
                        ConnectorSet pipe_connectorSet1 = pipe_connectorManager1.Connectors;
                        ConnectorManager pipe_connectorManager2 = pipes[i + 1].ConnectorManager;
                        ConnectorSet pipe_connectorSet2 = pipe_connectorManager2.Connectors;


                        // 2-3) �������� ������ �ִ� connectors ���� �����ų connector �̾Ƴ���

                        Connector pipe_connector1 = null;
                        Connector pipe_connector2 = null;
                        double minDist = double.MaxValue;
                        foreach (Connector connector1 in pipe_connectorSet1)
                        {
                            foreach (Connector connector2 in pipe_connectorSet2)
                            {
                                double d = connector1.Origin.DistanceTo(connector2.Origin);
                                if (d<minDist)
                                {
                                    pipe_connector1 = connector1;
                                    pipe_connector2 = connector2;
                                    minDist = d;
                                }
                            }
                        }

                        // 2-4) 2���� ������ �����Ű��
                        FamilyInstance fitting = document.Create.NewElbowFitting(pipe_connector1, pipe_connector2);
                    }

                    // (3) pipe�� ��ġ�� line �����ϱ�
                    foreach (Element element in elementlist)
                    {
                        ElementId elementId = element.Id;
                        document.Delete(elementId);
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Pipe ���� ������ �ʹ� ũ�ų� �۽��ϴ�. \n���� �� �ٽ� �õ����ּ���");
                    trans.RollBack();
                }

            }
        }

    }
}
