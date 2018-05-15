using ExecuteSOQL.SFDC;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Salesforce
{
    public class Execute_SOQL_PROD : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Username { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Password { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> SecurityToken { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> SOQLQuery { get; set; }

        [Category("Output")]
        public OutArgument<System.Data.DataTable> OutputDataTable { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            {
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string userName;
                string password;
                userName = Username.Get(context); //username from context
                password = Password.Get(context) + SecurityToken.Get(context);//password+token from context


                SforceService SfdcBinding = null;
                LoginResult CurrentLoginResult = null;

                SfdcBinding = new SforceService();
                try
                {
                    CurrentLoginResult = SfdcBinding.login(userName, password);

                }
                catch (System.Web.Services.Protocols.SoapException e)
                {
                    // This is likley to be caused by bad username or password
                    SfdcBinding = null;
                    throw (e);
                }
                catch (Exception e)
                {
                    // This is something else, probably comminication
                    SfdcBinding = null;
                    throw (e);
                }

                //Change the binding to the new endpoint
                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                Console.WriteLine(SfdcBinding.Url);
                Console.ReadLine();

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

                QueryResult queryResult = null;

                String SOQL = "";

                SOQL = SOQLQuery.Get(context); //SOQL Query from context

                queryResult = SfdcBinding.query(SOQL);

                System.Data.DataTable table = new System.Data.DataTable();

                if (queryResult.size > 0)
                {

                    int NoOfColumns = queryResult.records[0].Any.Count();
                    for (int i = 0; i < NoOfColumns; i++)
                    {
                        string ColumnName = queryResult.records[0].Any[i].Name;
                        table.Columns.Add(ColumnName);
                        //Console.WriteLine("Name of Column:" + ColumnName);
                    }

                    for (int j = 0; j < queryResult.records.Length; j++)
                    {

                        System.Data.DataRow dr = table.NewRow();
                        object[] rowArray = new object[table.Columns.Count];

                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            rowArray[k] = queryResult.records[j].Any[k].InnerText;
                        }
                        dr.ItemArray = rowArray;
                        table.Rows.Add(dr);
                    }

                }
                else
                {
                    Console.WriteLine("No Records Available.");
                }

                OutputDataTable.Set(context, table);//output table to context
            }
        }

    }
}
