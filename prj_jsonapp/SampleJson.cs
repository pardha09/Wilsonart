using System;
using System.Runtime.InteropServices;
using System.Net;

using System.Diagnostics;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Data;
using System.Web.Script.Serialization;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Linq;
namespace sample.programs
{
    class SampleJson
    {
        public void SayHello()
        {
            Console.WriteLine("Hello");
        }
        //This is  stub for CreateJson xxxxxxxxxxxx
        public bool CreateJSON(string pCompanyName, string pWarehouse, string pCustomerAccountNumber, string pShippingAddress, string pPurchaseOrder,
                string pOrderDate, string pInputFile, string pTargetFolder, string pDatacapID, string pUseQtyAsIs)
        {
            bool bRes = true;
            try
            {
                //Console.Write(DateTime.Now.ToString("MM/d/yyyy"));
                DataSet objDataSet = new DataSet();
                objDataSet.ReadXml(pInputFile, XmlReadMode.InferSchema);
              //  Console.Write(objDataSet.GetXml());
                DataTable dtFields = new DataTable();
                dtFields = objDataSet.Tables["F"];
                DataTable dtVariables = new DataTable();
                dtVariables = objDataSet.Tables["V"];
                DataRow[] FieldsRow = dtFields.Select("id LIKE 'LineItem%' and id <> 'LineItemFilterTerms'");

                //  Console.WriteLine("FieldsRow.Length--" + FieldsRow.Length);
                string sGlobalPartNumber = "";
                List<string> lstPartNumber = new List<string>();
                List<LineItem> openItems = new List<LineItem>();
                foreach (DataRow Row in FieldsRow)
                {

                    string sPartNumber = "";
                    int RowID = Convert.ToInt32(Row[0].ToString());
                    string ln_PartNumber = "";
                    string ln_Quantity = "";
                    string ln_PartNumber2 = "";
                    string ln_Quantity2 = "";
                    string ln_Price = "";
                    String ln_Amount, ln_Comments, ln_ExpDate, ln_Size, Ln_JobName, Ln_IsFSC = "";
                    // ArrayList lstPartNumber = new ArrayList();
                    if (pCompanyName.Equals("MIN_CF_NONSTOCKORDER") || pCompanyName.Equals("MIN_CF_MULTIUNITPROGRAM"))
                    {
                        //Begin
                        for (int i = RowID + 1; i <= RowID + 21; i++)
                        {
                            DataRow[] LineItemRows = dtVariables.Select("F_Id IN (" + i + ")");
                            foreach (DataRow LineItemRow in LineItemRows)
                            {
                                //Begin New code

                                if (i == RowID + 10)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        if(LineItemRow[1].ToString().Replace(":", "").Length<=2) sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "");
                                    }

                                }
                                else if (i == RowID + 11)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Replace(":", "").Length <= 3) sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "").Replace("1441", "144"); ;
                                    }

                                }
                                else if (i == RowID + 12)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Replace(":", "").Length <= 3) sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "").Replace("1441", "144");
                                    }

                                }
                                else if (i == RowID + 13)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 14)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Replace(":", "").Length <= 2) sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "");
                                    }

                                }
                                else if (i == RowID + 15)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Replace(":", "").Length == 3)
                                        {
                                            sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "") + "#";
                                        }
                                        else
                                        {
                                            sPartNumber = sPartNumber + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 16)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "");
                                    }

                                }
                                else if (i == RowID + 17)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "").Replace("1441", "144");
                                    }

                                }
                                else if (i == RowID + 18)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "").Replace("1441", "144");
                                    }

                                }
                                else if (i == RowID + 19)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "");
                                    }

                                }
                                else if (i == RowID + 20)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "");
                                    }

                                }
                                else if (i == RowID + 21)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        sPartNumber = sPartNumber + " " + LineItemRow[1].ToString().Replace(":", "");
                                    }

                                }
                            }

                        }
                        sPartNumber = sPartNumber.TrimStart().TrimEnd();
                        //Console.WriteLine("sPartNumber >> " + sPartNumber);
                        //if (sPartNumber.Split('#')[0].TrimStart().TrimEnd().Length > 10) { 
                        if (sPartNumber.Split('#')[0].TrimEnd().TrimStart().Split(' ').Length == 6) { 
                            ln_PartNumber = sPartNumber.Split('#')[0].Split(' ')[3] + ValidateFinish(sPartNumber.Split('#')[0].Split(' ')[4]) + sPartNumber.Split('#')[0].Split(' ')[5] + sPartNumber.Split('#')[0].Split(' ')[1] + sPartNumber.Split('#')[0].Split(' ')[2];
                            ln_PartNumber = ln_PartNumber.Replace(";", "");
                            ln_Quantity = sPartNumber.Split('#')[0].Split(' ')[0];
                        }
                                              

                        if (sGlobalPartNumber != ln_PartNumber)
                        {
                            if (ln_PartNumber.Length > 5)
                            {
                                var LineitemsObj1 = new LineItem
                                {
                                    PartNumber = ln_PartNumber.Replace(" ", string.Empty),
                                    Quantity = ln_Quantity.Replace(" ", string.Empty),
                                };

                                openItems.Add(LineitemsObj1);
                                Console.WriteLine(ln_Quantity  + " - " + ln_PartNumber);
                            }
                            sGlobalPartNumber = ln_PartNumber;

                            //   if (sPartNumber.Split('#')[1].TrimStart().TrimEnd().Length > 10)
                            if (sPartNumber.Split('#')[1].TrimEnd().TrimStart().Split(' ').Length ==6)
                            {
                                ln_PartNumber2 = sPartNumber.Split('#')[1].Split(' ')[4] + ValidateFinish(sPartNumber.Split('#')[1].Split(' ')[5]) + sPartNumber.Split('#')[1].Split(' ')[6] + sPartNumber.Split('#')[1].Split(' ')[2] + sPartNumber.Split('#')[1].Split(' ')[3];
                                ln_PartNumber2 = ln_PartNumber2.Replace(";", "");
                                ln_Quantity2 = sPartNumber.Split('#')[1].Split(' ')[1];
                                
                            }

                            if (ln_PartNumber2.Length > 5)
                            {
                                var LineitemsObj2 = new LineItem
                                {
                                    PartNumber = ln_PartNumber2.Replace(" ", string.Empty),
                                    Quantity = ln_Quantity2.Replace(" ", string.Empty),
                                };
                                openItems.Add(LineitemsObj2);
                                Console.WriteLine(ln_Quantity2  + " - " + ln_PartNumber2);
                            }


                          //  Console.WriteLine(" ln_PartNumber " + ln_PartNumber + " <<>> " + ln_PartNumber2);
                        }
                    }
//STOCK PO
                    else if (pCompanyName.Equals("MIN_CF_STOCKPO"))
                    {
                        //Begin
                        for (int i = RowID + 1; i <= RowID + 21; i++)
                        {
                            DataRow[] LineItemRows = dtVariables.Select("F_Id IN (" + i + ")");
                            foreach (DataRow LineItemRow in LineItemRows)
                            {
                                //Begin New code

                                if (i == RowID + 10)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_PartNumber = LineItemRow[1].ToString() + "#";
                                        
                                    }

                                }
                                else if (i == RowID + 11)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }
                                }
                                else if (i == RowID + 12)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }
                                }
                                else if (i == RowID + 13)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 14)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 15)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 16)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 17)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 18)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 19)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 20)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {

                                        if (LineItemRow[1].ToString().Length > 0)
                                        {
                                            ln_PartNumber = ln_PartNumber + LineItemRow[1].ToString() + "#";
                                        }
                                        else
                                        {
                                            ln_PartNumber = ln_PartNumber + "X" + "#";
                                        }
                                    }

                                }
                                else if (i == RowID + 21)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_PartNumber = ln_PartNumber + " " + LineItemRow[1];
                                    }

                                }
                            }

                        }
                        sPartNumber = ln_PartNumber.TrimStart().TrimEnd().Replace("|.","");
                        string[] sTemp = sPartNumber.Split('#');
                        if ((!sTemp[0].Contains('-')) && sTemp[sTemp.Length-1].Contains('-'))
                        {
                            sPartNumber = sTemp[sTemp.Length - 1].TrimStart().TrimEnd().Split(' ')[0] + sPartNumber;
                        }


                        sPartNumber = sPartNumber.Replace("#„#", "#X#").Replace("|,","");
                        if (ln_PartNumber.Contains("#"))
                        {
                            string sFinish = "";
                            string sJobName = "";
                            string sProductType = "";
                            string[] qtynumbers = sPartNumber.Split('#');
                            string sFinalPartNumber = "";
                            int iTemp = 0;
                            sJobName = qtynumbers[11];
 
                            for (int i=1;i<= qtynumbers.Length - 2; i++)
                            {
                                //Console.WriteLine(i + " - " + qtynumbers[i]);
                                
                                if (i==1 & int.TryParse(qtynumbers[i],out iTemp) & qtynumbers[0].Length>3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "3096";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] +" - "+sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                    //sFinalPartNumber = "";
                                }
                               else if (i == 2 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "30120";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                    //sFinalPartNumber = "";
                                }
                                else if (i == 3 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "30144";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                   // sFinalPartNumber = "";
                                }
                                else if (i == 4 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "36144";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                    //sFinalPartNumber = "";
                                }
                                else if (i == 5 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {

                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]); sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "4896";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                    //sFinalPartNumber = "";
                                }
                                else if (i == 6 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {

                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]); sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "48120";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                    //sFinalPartNumber = "";
                                }
                                else if (i == 7 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "48144";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                   // sFinalPartNumber = "";
                                }
                                else if (i == 8 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "6096";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                   // sFinalPartNumber = "";
                                }
                                else if (i == 9 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "60120";
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                   // sFinalPartNumber = "";
                                }
                                else if (i == 10 & int.TryParse(qtynumbers[i], out iTemp) & qtynumbers[0].Length > 3)
                                {
                                    sProductType = Get_ProductTypeForStandardFinish(qtynumbers[0].Split('-')[1]);
                                    sFinish = ValidateFinish(qtynumbers[0].Split('-')[1]);
                                    sFinalPartNumber = qtynumbers[0].Split('-')[0] + sFinish + sProductType + "60144" ;
                                    Console.WriteLine(sJobName + " - " + qtynumbers[i] + " - " + sFinalPartNumber);
                                    var LineitemsObj = new LineItem
                                    {
                                        PartNumber = sFinalPartNumber.Replace(" ", string.Empty),
                                        Quantity = qtynumbers[i].Replace(" ", string.Empty),
                                    };
                                    openItems.Add(LineitemsObj);
                                  
                                }
                              
                            }

                        }
                  
                    }
                    else
                    {
                        for (int i = RowID + 1; i <= RowID + 9; i++)
                        {
                            DataRow[] LineItemRows = dtVariables.Select("F_Id IN (" + i + ")");
                            foreach (DataRow LineItemRow in LineItemRows)
                            {
                                if (i == RowID + 1)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_PartNumber = LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 2)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_Quantity = LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 3)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_Price = LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 4)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_Amount = LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 5)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_Comments = LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 6)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_ExpDate = LineItemRow[1].ToString();
                                    }

                                }
                                else if (i == RowID + 7)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        ln_Size = LineItemRow[1].ToString();
                                    }

                                }

                                else if (i == RowID + 8)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        Ln_JobName = LineItemRow[1].ToString();
                                        // For Lake Area if there is a desciption with "Please" then add to main Node


                                    }

                                }
                                else if (i == RowID + 9)
                                {
                                    if (LineItemRow[0].ToString() == "FieldValue")
                                    {
                                        Ln_IsFSC = LineItemRow[1].ToString();
                                    }

                                }

                            }
                        }
                        
                        var LineitemsObj = new LineItem
                        {
                            PartNumber = ln_PartNumber.Replace(" ", string.Empty),
                            Quantity = ln_Quantity.Replace(" ", string.Empty),
                            //Price = ln_Price.Replace(" ", string.Empty),
                            //Amount = ln_Amount.Replace(" ", string.Empty),
                            //Comments = ln_Comments, 
                            // ExpDate = ln_ExpDate.Replace(" ", string.Empty),
                            //Size = ln_Size,
                            //Notes = Ln_JobName,
                            // Is_FSC = Ln_IsFSC
                        };
                        openItems.Add(LineitemsObj);
                    }
                }

                var obj = new POParams
                {

                    DatacapID = pDatacapID,
                    Warehouse = pWarehouse,
                    InforId = pCustomerAccountNumber,
                    Address = pShippingAddress,
                    PurchaseOrder = pPurchaseOrder,
                    OrderDate = pOrderDate,
                    UseQuantityAsIs = pUseQtyAsIs,
                    LineItem = openItems.GroupBy(k => k.PartNumber).Select(g => g.First()).ToList()
                };

                var DList = openItems.GroupBy(k => k.PartNumber).Select(g => g.First()).ToList();
                Console.WriteLine("openItems List Count " + DList.Count());


                
                var json = new JavaScriptSerializer().Serialize(obj);
                File.WriteAllText(pTargetFolder + "metadata_" + pPurchaseOrder + ".json", json);
                bRes = true;

                // Console.WriteLine(DList.Count);
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                // It is a best practice to have a try catch in every action to prevent any unexpected errors
                // from being thrown back to RRS.
                Console.WriteLine("There was an exception: " + ex.Message);
                Console.ReadKey();
                bRes = false;
            }

            return bRes;
        }

        //Checks for proper Finish format
        public String ValidateFinish(string Finish)
        {

            try
            {
                if (Finish.TrimEnd().TrimStart().Length == 0) return "";

                Finish = Finish.ToUpper();
                if (Finish.ToUpper().Contains("K"))
                {
                    if (Finish.ToUpper().Substring(0, 1).Equals("K"))
                    {
                        //In case Standard finishes 38 and 60 are passed with "K" before
                        if (Finish.Substring(1, 2).Equals("38") || Finish.Substring(1, 2).Equals("60"))
                        {
                            Finish = Finish.Replace("K", "");
                            Finish = Finish.Replace("K", "");
                            return Finish;
                        }
                    }
                    //In case Standard finishes 38 and 60 are passed with "K" after
                    if (Finish.ToUpper().Substring(2, 1).Equals("K"))
                    {
                        if (Finish.Substring(0, 2).Equals("38") || Finish.Substring(0, 2).Equals("60"))
                        {
                            Finish = Finish.Replace("K", "");
                            Finish = Finish.Replace("k", "");
                            return Finish;
                        }

                    }

                    int iPostionOfK = Finish.ToUpper().IndexOf("K");
                    //In case Finish is not in the beginning then insert K at the beginning by repositioning
                    if (iPostionOfK != 0)
                    {
                        Finish = Finish.Replace("K", "");
                        Finish = Finish.Insert(0, "K");
                    }
                    //Incase Finish is having "0" when K is there then remove "0" eg. K07 will be replaced like K7
                    if (Finish.ToUpper().Contains("K"))
                    {
                        if (Finish.Substring(1, 1).Equals("0"))
                        {
                            Finish = Finish.Replace("K0", "K");
                        }
                    }

                    // Console.WriteLine(Finish);
                    // Console.ReadLine();
                }
                else
                {
                    if (Finish != "38" && Finish != "60")
                    {
                        Finish = "K" + Finish;
                        if (Finish.Substring(1, 1).Equals("0"))
                        {
                            Finish = Finish.Replace("K0", "K");
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occured: " + ex.ToString());
            }

            return Finish;
        }
// Get Product type by Finish Value
        public String Get_ProductTypeForStandardFinish(string sFinish)
        {
            String sProductType = "";
            try
            {
                if (sFinish.Equals("35") || sFinish.Equals("45") || sFinish.Equals("55") || sFinish.Equals("57"))
                {

                    sProductType = "376";
                }
                else
                {
                    sProductType = "350";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occured: " + ex.ToString());
            }

            return sProductType;
        }

        public class POParams
        {
            public string DatacapID { get; set; }
            public string Warehouse { get; set; }
            public string InforId { get; set; }
            public string Address { get; set; }
            public string PurchaseOrder { get; set; }
            public string OrderDate { get; set; }
            public string UseQuantityAsIs { get; set; }
            public List<LineItem> LineItem { get; set; }
        }

        public class LineItem
        {
            public string PartNumber { get; set; }
            public string Quantity { get; set; }
        }
    }


}
