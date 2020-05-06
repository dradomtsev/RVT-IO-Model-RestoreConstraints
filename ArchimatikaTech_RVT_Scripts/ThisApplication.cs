/*
 * Created by SharpDevelop.
 * User: d.radomtsev
 * Date: 27.12.2018
 * Time: 13:03
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BoundarySegment = Autodesk.Revit.DB.BoundarySegment;

namespace ArchimatikaTech_RVT_Scripts
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("12EF2991-8B35-46A6-A203-288CBAC07038")]
	public partial class ThisApplication
	{
        public void SetSpacesConstraintsfromLink()
        {
            //Allocate resources
            StringBuilder stServiceString = new StringBuilder();
            IList<Parameter> prSpace = new List<Parameter>();
            IList<Element> linkedRooms = new List<Element>();
            Element RVTLink;
            RevitLinkType linkType;
            double dResult = 0;
            const double dMetertoFeet = 304.8;
            StringBuilder sdResult = new StringBuilder();
            TaskDialog tskDiag = new TaskDialog("SetSpacesConstraints()");
            //tskDiag.AllowCancellation = false;
            // All actions wraps into transactoions
            using (Transaction trans = new Transaction(this.ActiveUIDocument.Document, "SetSpacesConstraints"))
            {
                // Start of transaction
                try
                {
                    trans.Start();
                    //tskDiag.MainInstruction = "Start of SetSpacesConstraints()";
                    //TaskDialogResult tResult = tskDiag.Show();
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("Transaction", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
                finally
                {
                    stServiceString.Clear();
                }

                Document curDoc = this.ActiveUIDocument.Document;
                FilteredElementCollector collector = new FilteredElementCollector(curDoc);
                try
                {
                    //tskDiag.MainInstruction = "Get all linked files";
                    //TaskDialogResult tResult = tskDiag.Show();

                    RVTLink = collector.OfCategory(BuiltInCategory.OST_RvtLinks).OfClass(typeof(RevitLinkType)).Where(psp => psp.Name.Contains("Zones")).FirstOrDefault();
                    linkType = RVTLink as RevitLinkType;
                    foreach (Document linkedDoc in this.Application.Documents)
                    {
                        IList<String> st = new List<String>();
                        st = linkType.Name.Split('.');
                        if (linkedDoc.Title.Equals(st.First()))
                        {
                            using (FilteredElementCollector collLinked = new FilteredElementCollector(linkedDoc))
                            {
                                linkedRooms = collLinked.OfCategory(BuiltInCategory.OST_Rooms).ToElements();
                            }
                        }
                    }
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("Transaction", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
                finally
                {
                    stServiceString.Clear();
                }

                // Get Lists of rooms
                IList<Room> rooms = new List<Room>(linkedRooms.Cast<Room>());
                // Get Lists of spaces
                using (FilteredElementCollector spaces = new FilteredElementCollector(curDoc))
                {
                    try
                    {
                        spaces.WherePasses(new SpaceFilter());

                        if (rooms.Count != 0 & spaces.IsValidObject)
                        {
                            //Space spItem = spaces.Where(psp => psp.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() == "1201-1").Cast<Space>().FirstOrDefault();
                            foreach (Space spItem in spaces)
                            {
                                StringBuilder sttemp = new StringBuilder();
                                sttemp.Append(spItem.get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NUMBER).AsString());
                                Room rmItem = rooms.Where(psp => psp.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() == sttemp.ToString()).Cast<Room>().FirstOrDefault();
                                try
                                {
                                    prSpace = spItem.GetParameters("Base Offset");
                                    dResult = rmItem.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).AsDouble();
                                    prSpace.First().Set(dResult);// * dMetertoFeet);
                                    prSpace.Clear();
                                }
                                catch (NullReferenceException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (InvalidOperationException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (ArgumentException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }

                                try
                                {
                                    Level Newlev = (from el in new FilteredElementCollector(curDoc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>()
                                                    where el.Name == rmItem.UpperLimit.Name
                                                    select el).FirstOrDefault();
                                    prSpace = spItem.GetParameters("Upper Limit");
                                    prSpace.First().Set(Newlev.Id);
                                    sdResult.Clear();
                                    prSpace.Clear();
                                }
                                catch (NullReferenceException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (InvalidOperationException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (ArgumentException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }

                                try
                                {
                                    prSpace = spItem.GetParameters("Limit Offset");
                                    dResult = rmItem.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble();
                                    prSpace.First().Set(dResult);// * dMetertoFeet);
                                    prSpace.Clear();
                                }
                                catch (NullReferenceException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (InvalidOperationException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (ArgumentException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                //try
                                //{
                                //    stServiceString.Clear();
                                //    stServiceString.AppendLine("Space Name: " + spItem.Name);
                                //    stServiceString.AppendLine("Space Base level: " + spItem.Level.Name);
                                //    stServiceString.AppendLine("Space Base offset: " + spItem.BaseOffset);
                                //    stServiceString.AppendLine("Space Upper Level: " + spItem.UpperLimit.Name);
                                //    stServiceString.AppendLine("Space Upper offset: " + spItem.LimitOffset);
                                //    tskDiag.MainInstruction = stServiceString.ToString();
                                //    TaskDialogResult tskresult = tskDiag.Show();
                                //}
                                //catch (Autodesk.Revit.Exceptions.InternalException e)
                                //{
                                //    stServiceString.AppendLine(e.Message);
                                //    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                //    stServiceString.Clear();
                                //}
                                sttemp.Clear();
                            }
                        }
                        else
                            throw new ArgumentNullException();
                    }
                    catch (ArgumentNullException e)
                    {
                        stServiceString.AppendLine(e.Message);
                        TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                        stServiceString.Clear();
                    }
                    finally
                    {
                        stServiceString.Clear();
                    }
                }
                    
                // Commit of transaction
                try
                {   
                    trans.Commit();
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("Transaction", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
                finally
                {
                    stServiceString.Clear();
                }

                // Release resources
                try
                {
                    prSpace.Clear();
                    stServiceString.Clear();
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
            }
        }

        public void SetSpacesConstraints()
        {
            //Allocate resources
            StringBuilder stServiceString = new StringBuilder();
            IList<Parameter> prSpace = new List<Parameter>();
            IList<Element> linkedRooms = new List<Element>();
            Element RVTLink;
            RevitLinkType linkType;
            double dResult = 0;
            const double dMetertoFeet = 304.8;
            StringBuilder sdResult = new StringBuilder();
            TaskDialog tskDiag = new TaskDialog("SetSpacesConstraints()");
            //tskDiag.AllowCancellation = false;
            // All actions wraps into transactoions
            using (Transaction trans = new Transaction(this.ActiveUIDocument.Document, "SetSpacesConstraints"))
            {
                // Start of transaction
                try
                {
                    trans.Start();
                    //tskDiag.MainInstruction = "Start of SetSpacesConstraints()";
                    //TaskDialogResult tResult = tskDiag.Show();
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("Transaction", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
                finally
                {
                    stServiceString.Clear();
                }

                Document curDoc = this.ActiveUIDocument.Document;
                FilteredElementCollector collector = new FilteredElementCollector(curDoc);
                //try
                //{
                //    //tskDiag.MainInstruction = "Get all linked files";
                //    //TaskDialogResult tResult = tskDiag.Show();

                //    RVTLink = collector.OfCategory(BuiltInCategory.OST_RvtLinks).OfClass(typeof(RevitLinkType)).Where(psp => psp.Name.Contains("Zones")).FirstOrDefault();
                //    linkType = RVTLink as RevitLinkType;
                //    foreach (Document linkedDoc in this.Application.Documents)
                //    {
                //        IList<String> st = new List<String>();
                //        st = linkType.Name.Split('.');
                //        if (linkedDoc.Title.Equals(st.First()))
                //        {
                //            using (FilteredElementCollector collLinked = new FilteredElementCollector(linkedDoc))
                //            {
                //                linkedRooms = collLinked.OfCategory(BuiltInCategory.OST_Rooms).ToElements();
                //            }
                //        }
                //    }
                //}
                //catch (InvalidOperationException e)
                //{
                //    stServiceString.AppendLine(e.Message);
                //    TaskDialog.Show("Transaction", "Error: " + stServiceString.ToString());
                //    stServiceString.Clear();
                //}
                //finally
                //{
                //    stServiceString.Clear();
                //}

                // Get Lists of rooms
                FilteredElementCollector rooms = new FilteredElementCollector(curDoc);
                rooms.WherePasses(new RoomFilter());
                // Get Lists of spaces
                using (FilteredElementCollector spaces = new FilteredElementCollector(curDoc))
                {
                    try
                    {
                        spaces.WherePasses(new SpaceFilter());

                        if (spaces.IsValidObject)
                        {
                            //Space spItem = spaces.Where(psp => psp.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() == "1201-1").Cast<Space>().FirstOrDefault();
                            foreach (Space spItem in spaces)
                            {
                                StringBuilder sttemp = new StringBuilder();
                                sttemp.Append(spItem.get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NUMBER).AsString());
                                Room rmItem = rooms.Where(psp => psp.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() == sttemp.ToString()).Cast<Room>().FirstOrDefault();
                                try
                                {
                                    prSpace = spItem.GetParameters("Base Offset");
                                    dResult = rmItem.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).AsDouble();
                                    prSpace.First().Set(dResult);// * dMetertoFeet);
                                    prSpace.Clear();
                                }
                                catch (NullReferenceException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (InvalidOperationException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (ArgumentException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }

                                try
                                {
                                    Level Newlev = (from el in new FilteredElementCollector(curDoc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>()
                                                    where el.Name == rmItem.UpperLimit.Name
                                                    select el).FirstOrDefault();
                                    prSpace = spItem.GetParameters("Upper Limit");
                                    prSpace.First().Set(Newlev.Id);
                                    sdResult.Clear();
                                    prSpace.Clear();
                                }
                                catch (NullReferenceException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (InvalidOperationException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (ArgumentException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }

                                try
                                {
                                    prSpace = spItem.GetParameters("Limit Offset");
                                    dResult = rmItem.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble();
                                    prSpace.First().Set(dResult);// * dMetertoFeet);
                                    prSpace.Clear();
                                }
                                catch (NullReferenceException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (InvalidOperationException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                catch (ArgumentException e)
                                {
                                    stServiceString.AppendLine(e.Message);
                                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                    stServiceString.Clear();
                                }
                                //try
                                //{
                                //    stServiceString.Clear();
                                //    stServiceString.AppendLine("Space Name: " + spItem.Name);
                                //    stServiceString.AppendLine("Space Base level: " + spItem.Level.Name);
                                //    stServiceString.AppendLine("Space Base offset: " + spItem.BaseOffset);
                                //    stServiceString.AppendLine("Space Upper Level: " + spItem.UpperLimit.Name);
                                //    stServiceString.AppendLine("Space Upper offset: " + spItem.LimitOffset);
                                //    tskDiag.MainInstruction = stServiceString.ToString();
                                //    TaskDialogResult tskresult = tskDiag.Show();
                                //}
                                //catch (Autodesk.Revit.Exceptions.InternalException e)
                                //{
                                //    stServiceString.AppendLine(e.Message);
                                //    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                                //    stServiceString.Clear();
                                //}
                                sttemp.Clear();
                            }
                        }
                        else
                            throw new ArgumentNullException();
                    }
                    catch (ArgumentNullException e)
                    {
                        stServiceString.AppendLine(e.Message);
                        TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                        stServiceString.Clear();
                    }
                    finally
                    {
                        stServiceString.Clear();
                    }
                }

                // Commit of transaction
                try
                {
                    trans.Commit();
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("Transaction", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
                finally
                {
                    stServiceString.Clear();
                }

                // Release resources
                try
                {
                    prSpace.Clear();
                    stServiceString.Clear();
                }
                catch (InvalidOperationException e)
                {
                    stServiceString.AppendLine(e.Message);
                    TaskDialog.Show("SetSpacesConstraints", "Error: " + stServiceString.ToString());
                    stServiceString.Clear();
                }
            }
        }

        public void SetRoomsConstraints()
        {
            //String stPrompt = "Show parameters in selected Element: \n\rmItem";
            StringBuilder st = new StringBuilder();
            IList<Parameter> prRoom = new List<Parameter>();
            //IList<Parameter> prSpace = new List<Parameter>();
            double dResult = 0;
            const double dMetertoFeet = 0.3048;
            StringBuilder sdResult = new StringBuilder();

            // All actions wraps into transactoions
            using (Transaction trans = new Transaction(this.ActiveUIDocument.Document, "SetRoomsConstraints"))
            {
                // Start of transaction
                trans.Start();
                Document curDoc = this.ActiveUIDocument.Document;
                // Get Lists of rooms and spaces
                FilteredElementCollector rooms = new FilteredElementCollector(curDoc);
                rooms.WherePasses(new RoomFilter());
                //FilteredElementCollector spaces = new FilteredElementCollector(curDoc);
                //spaces.WherePasses(new SpaceFilter());
                // Check valid display units
                //IList<DisplayUnitType> disp = UnitUtils.GetValidDisplayUnits();
                //foreach(DisplayUnitType d in disp) {st.AppendLine(d.ToString());}
                //TaskDialog.Show("Revit", stPrompt  + st.ToString());

                if (rooms.IsValidObject)
                {
                    //foreach (Space spItem in spaces)
                    //{
                    //    foreach (Parameter para in spaces.First().Parameters)
                    //    { st.AppendLine(GetParameterInformation(para, curDoc)); }
                    //    TaskDialog.Show("Revit", stPrompt + st.ToString());
                    //    st.Clear();
                    //}

                    foreach (Room rmItem in rooms)
                    {
                        if (rmItem.IsValidObject)
                        {
                            //Space sptemp = (from tSp in spaces.Cast<Space>() where tSp.Name == rmItem.Name select tSp).FirstOrDefault();
                            prRoom = rmItem.GetParameters("Base Offset (Pset_SpaceConstrains)");
                            if (prRoom.First().HasValue == true)
                            {
                                dResult = prRoom.First().AsDouble();
                                prRoom = rmItem.GetParameters("Base Offset");
                                //prSpace = spaces.Where(psp => psp.GetParameters("Room Number").FirstOrDefault().AsString() == rmItem.Number).Cast<Space>().FirstOrDefault().GetParameters("Base Offset");
                                prRoom.First().Set(dResult / dMetertoFeet);
                                //prSpace.First().Set(dResult / dMetertoFeet);
                                dResult = 0;
                                //st.Clear();							
                            }

                            prRoom = rmItem.GetParameters("Upper Level (Pset_SpaceConstrains)");
                            if (prRoom.First().HasValue == true)
                            {
                                sdResult.Append(prRoom.First().AsString());
                                Level Newlev = (from el in new FilteredElementCollector(curDoc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>()
                                                where el.Name == sdResult.ToString()
                                                select el).FirstOrDefault();
                                prRoom = rmItem.GetParameters("Upper Limit");
                                //prSpace = spaces.Where(psp => psp.GetParameters("Room Number").FirstOrDefault().AsString() == rmItem.Number).Cast<Space>().FirstOrDefault().GetParameters("Upper Limit");
                                prRoom.First().Set(Newlev.Id);
                                //prSpace.First().Set(Newlev.Id);
                                sdResult.Clear();
                            }

                            prRoom = rmItem.GetParameters("Upper Offset (Pset_SpaceConstrains)");
                            if (prRoom.First().HasValue == true)
                            {
                                dResult = prRoom.First().AsDouble();
                                prRoom = rmItem.GetParameters("Limit Offset");
                                //prSpace = spaces.Where(psp => psp.GetParameters("Room Number").FirstOrDefault().AsString() == rmItem.Number).Cast<Space>().FirstOrDefault().GetParameters("Limit Offset");
                                prRoom.First().Set(dResult / dMetertoFeet);
                                //prSpace.First().Set(dResult / dMetertoFeet);
                                dResult = 0;
                                sdResult.Clear();
                            }
                        }
                    }
                }
                trans.Commit();
            }
        }

        public String GetParameterInformation(Parameter para, Document document)
        {
            string defName = para.Definition.Name + "\t : ";
            string defValue = string.Empty;
            // Use different method to get parameter data according to the storage type
            switch (para.StorageType)
            {
                case StorageType.Double:
                    //covert the number into Metric
                    defValue = para.AsValueString();
                    break;
                case StorageType.ElementId:
                    //find out the name of the element
                    Autodesk.Revit.DB.ElementId id = para.AsElementId();
                    if (id.IntegerValue >= 0)
                    {
                        defValue = document.GetElement(id).Name;
                    }
                    else
                    {
                        defValue = id.IntegerValue.ToString();
                    }
                    break;
                case StorageType.Integer:
                    if (ParameterType.YesNo == para.Definition.ParameterType)
                    {
                        if (para.AsInteger() == 0)
                        {
                            defValue = "False";
                        }
                        else
                        {
                            defValue = "True";
                        }
                    }
                    else
                    {
                        defValue = para.AsInteger().ToString();
                    }
                    break;
                case StorageType.String:
                    defValue = para.AsString();
                    break;
                default:
                    defValue = "Unexposed parameter.";
                    break;
            }

            return defName + defValue;
        }

        private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
	}
}