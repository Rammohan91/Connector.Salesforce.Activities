using ExecuteSOQL.SFDC_TEST;
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
    public class Upsert_Record_SANDBOX : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Username { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Password { get; set; }

        [Category("Input")]
        public InArgument<String> SecurityToken { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String[]> FieldNames { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String[]> FieldValues { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> ExternalIDFieldName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> ObjectName { get; set; }


        [Category("Output")]
        public OutArgument<String> RecordID { get; set; }

        protected override void Execute(CodeActivityContext context)
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

            //Console.WriteLine(SfdcBinding.Url);
            //Console.ReadLine();

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

            String[] fieldNames = FieldNames.Get(context);
            String[] fieldValues = FieldValues.Get(context);

            sObject obj = new sObject();
            System.Xml.XmlElement[] objFields = new System.Xml.XmlElement[fieldNames.Length];

            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();


            for (int i = 0; i < fieldNames.Length; i++)
            {
                objFields[i] = doc.CreateElement(fieldNames[i]);
            }

            for (int j = 0; j < fieldValues.Length; j++)
            {
                objFields[j].InnerText = fieldValues[j];
            }

            obj.type = ObjectName.Get(context);
            obj.Any = objFields;

            sObject[] objList = new sObject[1];
            objList[0] = obj;

            UpsertResult[] results = SfdcBinding.upsert(ExternalIDFieldName.Get(context), objList);


            for (int j = 0; j < results.Length; j++)
            {
                if (results[j].success)
                {
                    RecordID.Set(context, "Record Created/Updated for ID: " + results[j].id);
                }
                else
                {
                    // There were errors during the create call,
                    // go through the errors array and write
                    // them to the console
                    String error;
                    for (int i = 0; i < results[j].errors.Length; i++)
                    {
                        Error err = results[j].errors[i];
                        error = "Errors was found on item " + j.ToString() + Environment.NewLine
                            + "Error code is: " + err.statusCode.ToString() + Environment.NewLine
                            + "Error message: " + err.message;
                        RecordID.Set(context, error);
                    }
                }
            }
        }

    }

}
