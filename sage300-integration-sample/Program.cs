using AccpacCOMAPI;
using Sage300IntegrationSample.Models;
//******************************************************
//*
//* Written by Branko Pedisic // Ardento Pty Ltd
//* 
//* THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
//* ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
//* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
//* IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
//* INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT 
//* NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR 
//* PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
//* WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
//* ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY 
//* OF SUCH DAMAGE. 
//* 
//******************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Sage300IntegrationSample
{
    class Program
    {
        static void Main(string[] args)
        {
            UploadAPInvoicesToSage300();

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Press any key to exit");

            while (!Console.KeyAvailable)
            {
                // wait for the "any" key
                // before exiting.
            }
        }

        private static void UploadAPInvoicesToSage300()
        {
            // Sage 300 session object.
            AccpacSession session = null;

            // Sage 300 db link object.
            AccpacDBLink dbLink = null;

            // Sage 300 view declarations for AP invoice
            // batch/entry.
            AccpacView vAPIBC = null;   // AP0020
            AccpacView vAPIBH = null;   // AP0021
            AccpacView vAPIBD = null;   // AP0022
            AccpacView vAPIBS = null;   // AP0023
            AccpacView vAPIBHO = null;  // AP0402
            AccpacView vAPIBDO = null;  // AP0401

            try
            {
                // Populate our data model with data read from
                // our test XML data file.
                var invoices = ReadData();

                if (invoices.Count == 0)
                {
                    // If no invoices were found, then we'll show the message
                    // in the console window and exit this function.
                    Console.WriteLine("No invoices found.");
                    return;
                }

                Console.WriteLine(invoices.Count + " invoices found. Now attempting to process into "
                    + "Sage...");

                // Initialise and open session to Sage 300.
                // Hardcoding the default sample company user "ADMIN"
                // and sample company id "SAMLTD".
                // We could open up a signon manager window here if we wanted
                // the user to specify logon credentials at runtime.
                #region Session.
                session = new AccpacSession();
                session.Init("", "AS", "AS1001", "62A");
                session.Open("ADMIN", "ADMIN", "SAMLTD", System.DateTime.Now, 0, "");
                dbLink = session.OpenDBLink(tagDBLinkTypeEnum.DBLINK_COMPANY, tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE);
                Console.WriteLine("Session and db link created.");
                #endregion Session.

                // Write to Sage 300.
                #region Post to Sage.

                // Open views.
                dbLink.OpenView("AP0020", out vAPIBC);
                dbLink.OpenView("AP0021", out vAPIBH);
                dbLink.OpenView("AP0022", out vAPIBD);
                dbLink.OpenView("AP0023", out vAPIBS);
                dbLink.OpenView("AP0402", out vAPIBHO);
                dbLink.OpenView("AP0401", out vAPIBDO);

                // Compose views.
                vAPIBC.Compose(new AccpacView[] { vAPIBH });
                vAPIBH.Compose(new AccpacView[] { vAPIBC, vAPIBD, vAPIBS, vAPIBHO });
                vAPIBD.Compose(new AccpacView[] { vAPIBH, vAPIBC, vAPIBDO });
                vAPIBS.Compose(new AccpacView[] { vAPIBH });
                vAPIBHO.Compose(new AccpacView[] { vAPIBH });
                vAPIBDO.Compose(new AccpacView[] { vAPIBD });

                // Create new batch.
                vAPIBC.RecordCreate(tagViewRecordCreateEnum.VIEW_RECORD_CREATE_INSERT);

                // Create new batch entry for each invoice read from the 
                // csv file.
                foreach (var invoice in invoices)
                {
                    // Invoice header.
                    vAPIBH.RecordCreate(tagViewRecordCreateEnum.VIEW_RECORD_CREATE_NOINSERT);
                    vAPIBH.Fields.get_FieldByName("IDVEND").set_Value(invoice.VendorID);
                    vAPIBH.Fields.get_FieldByName("INVCDESC").set_Value(invoice.Description);
                    vAPIBH.Fields.get_FieldByName("IDINVC").set_Value(invoice.InvoiceNumber);
                    vAPIBH.Fields.get_FieldByName("DATEINVC").set_Value(invoice.InvoiceDate);
                    vAPIBH.Fields.get_FieldByName("AMTGROSTOT").set_Value(invoice.DocumentTotal);
                    vAPIBH.Fields.get_FieldByName("TEXTTRX").set_Value("1"); // Invoice.

                    // Add invoice details.
                    foreach (var line in invoice.Lines)
                    {
                        vAPIBD.RecordCreate(tagViewRecordCreateEnum.VIEW_RECORD_CREATE_NOINSERT);
                        vAPIBD.Fields.get_FieldByName("TEXTDESC").set_Value(line.Description);
                        vAPIBD.Fields.get_FieldByName("IDGLACCT").set_Value(line.AccountCode);
                        vAPIBD.Fields.get_FieldByName("AMTDIST").set_Value(line.AmountEx);
                        vAPIBD.Insert();
                    }

                    // Insert header to the batch.
                    vAPIBH.Insert();

                    Console.WriteLine("..invoice \"" + invoice.InvoiceNumber + "\" has been added to the batch.");
                }

                // Post the batch.
                vAPIBC.Fields.get_FieldByName("BTCHDESC").set_Value("Test batch import from .NET sample - " + System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"));
                vAPIBC.Update();

                Console.WriteLine("Batch posted into Sage. " 
                    + invoices.Count + " invoices have been posted to batch number "
                    + vAPIBC.Fields.get_FieldByName("CNTBTCH").get_Value().ToString() + ".");

                #endregion Post to Sage.
            }
            catch (Exception ex)
            {
                if (session != null)
                {
                    for (int i = 0; i < session.Errors.Count; i++)
                    {
                        Console.WriteLine(session.Errors.Item(i));
                    }
                }
                else
                {
                    Console.WriteLine(ex.Message);

                    if (ex.InnerException != null)
                        Console.WriteLine(ex.InnerException.Message);
                }
            }
            finally
            {
                // Cleanup any open views, db links
                // and sessions.
                if (vAPIBC != null)
                    vAPIBC.Close();

                if (vAPIBH != null)
                    vAPIBH.Close();

                if (vAPIBD != null)
                    vAPIBD.Close();

                if (vAPIBS != null)
                    vAPIBS.Close();

                if (vAPIBHO != null)
                    vAPIBHO.Close();

                if (vAPIBDO != null)
                    vAPIBDO.Close();

                if (dbLink != null)
                    dbLink.Close();

                if (session != null)
                    session.Close();
            }
        }

        private static List<InvoiceModel> ReadData()
        {
            var invoices = new List<InvoiceModel>();

            try
            {
                XDocument doc = XDocument.Load(@"..\..\Resources\Data\TestData.xml");

                invoices = doc.Root
                    .Elements("Invoice")
                    .Select(x => new InvoiceModel
                    {
                        VendorID = (string)x.Element("VendorID"),
                        Description = (string)x.Element("Description"),
                        InvoiceNumber = (string)x.Element("InvoiceNumber"),
                        InvoiceDate = DateTime.Parse((string)x.Element("InvoiceDate")),
                        DocumentTotal = double.Parse((string)x.Element("DocumentTotal")),
                        Lines = x.Descendants("Line")
                          .Select(d => new InvoiceLineModel
                          {
                              Description = (string)d.Element("Description"),
                              AccountCode = (string)d.Element("AccountCode"),
                              AmountEx = double.Parse((string)d.Element("AmountEx"))
                          }).ToList()
                    })
                    .ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in read data. " + ex.Message);
            }

            return invoices;
        }
    }
}
