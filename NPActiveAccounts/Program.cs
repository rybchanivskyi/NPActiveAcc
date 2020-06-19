using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPActiveAccount
{
    class Program
    {
        static void Main(string[] args)
        {
            const string url = "";
            const string userName = "";
            const string password = "";
            const string fileName = "ActiveAccounts.txt";

            string conn = $@"
            Url = {url};
            AuthType = Office365;
            UserName = {userName};
            Password = {password};
            RequireNewInstance = True";
            try
            {
                Executor(conn, fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: \n" + ex.Message);
            }
            Console.WriteLine("Press any key to exit.");
            Console.ReadLine();
        }

        public static void WriteRecord(string companyId, StreamWriter file)
        {
            file.WriteLine(companyId);
        }

        public static List<Entity> getDataOfQuery(CrmServiceClient svc, QueryExpression query)
        {
            int pageNumber = 1;
            query.PageInfo = new PagingInfo();
            query.PageInfo.PageNumber = pageNumber;
            query.PageInfo.PagingCookie = null;
            List<Entity> result = new List<Entity>();
            while (true)
            {
                EntityCollection pageRecords = svc.RetrieveMultiple(query);
                if (pageRecords.Entities != null)
                {
                    result.AddRange(pageRecords.Entities.ToList());
                }

                if (pageRecords.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = pageRecords.PagingCookie;
                }
                else
                {
                    break;
                }
            }
            return result;
        }

        public static void Executor(string connectionStr, string fileName)
        {
            using (var svc = new CrmServiceClient(connectionStr))
            {
                Console.WriteLine("Try to receive data!");


                using (StreamWriter file = new StreamWriter(@fileName, false))
                {
                    Console.WriteLine($@"Start creating records in file: {fileName}");

                    QueryExpression query = new QueryExpression()
                    {
                        Distinct = false,
                        EntityName = "account",
                        ColumnSet = new ColumnSet(new[] { "accountid", "emailaddress1" }),
                        Criteria =
                        {
                            Filters =
                            {
                                new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions =
                                    {
                                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                                        new ConditionExpression("emailaddress1", ConditionOperator.NotEqual, String.Empty)
                                    }
                                }
                            }
                        }
                    };
                    QueryExpression queryContacts = new QueryExpression()
                    {
                        Distinct = false,
                        EntityName = "contact",
                        ColumnSet = new ColumnSet(new[] { "contactid", "accountid", "emailaddress1" }),
                        Criteria =
                        {
                            Filters =
                            {
                                new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions =
                                    {
                                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                                        new ConditionExpression("emailaddress1", ConditionOperator.NotEqual, String.Empty)
                                    }
                                }
                            }
                        }
                    };
                    List<Entity> accounts = getDataOfQuery(svc, query);
                    List<Entity> contacts = getDataOfQuery(svc, queryContacts);

                    foreach (Entity account in accounts)
                    {
                        bool hasEmail = false;
                        foreach (Entity contact in contacts)
                        {
                            try
                            {
                                if (account.Attributes["accountid"].ToString() ==
                                    contact.Attributes["accountid"].ToString()
                                    && account.Attributes["emailaddress1"].ToString() ==
                                    contact.Attributes["emailaddress1"].ToString())
                                {
                                    hasEmail = true;
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("An error occurred: \n" + ex.Message);
                            }
                        }
                        if (!hasEmail)
                        {
                            WriteRecord(account["accountid"].ToString(), file);
                        }
                    }
                }
                Console.WriteLine($@"All data was written to file - {Environment.CurrentDirectory}\{fileName}");
            }
        }
    }
}
