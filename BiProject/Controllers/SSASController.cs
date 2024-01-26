using Microsoft.AspNetCore.Mvc;
using Microsoft.AnalysisServices.AdomdClient;
using Newtonsoft.Json;
using System.Data.Common;

namespace BiProject.Controllers
{
        public class SSASController : Controller

        {
        private static readonly string CNX_STRING = "Provider=MSOLAP; Data Source=CFEKI_W_P;Initial Catalog=Comms_DW;";
        private static readonly string CUBE = "[Comms DW]";
            private readonly IConfiguration _configuration;

            public SSASController(IConfiguration configuration)
            {
                _configuration = configuration;
            }
        /**
         * Simple queries by a single factor 
         */
        [Route("api/getByCampaign")]
        public IActionResult GetByCampaign(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Campaigns].[Campaign Name].[Campaign Name]" };
            string connectionString = CNX_STRING;


            System.Text.StringBuilder result = new System.Text.StringBuilder();



            //foreach (var measure in split_measures)
            //{
            using (AdomdConnection conn = new AdomdConnection(connectionString))
            {
                conn.Open();

                using (AdomdCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = BuildCommand(split_measures,dimensions);
                    Console.WriteLine(cmd.CommandText);

                    CellSet cs = cmd.ExecuteCellSet();

                    foreach (var dim in dimensions) {
                        result.Append(",");
                    }

                    TupleCollection tuplesOnColumns = cs.Axes[0].Set.Tuples;
                    foreach (Microsoft.AnalysisServices.AdomdClient.Tuple column in tuplesOnColumns)
                    {
                        result.Append(column.Members[0].Caption + ",");
                    }
                    result.Length--;
                    result.Append(";");
                    
                    //Output the row captions from the second axis and cell data
                    //Note that this procedure assumes a two-dimensional cellset
                    TupleCollection tuplesOnRows = cs.Axes[1].Set.Tuples;
                    for (int row = 0; row < tuplesOnRows.Count; row++)
                    {
                        result.Append(tuplesOnRows[row].Members[0].Caption + ",");
                        for (int col = 0; col < tuplesOnColumns.Count; col++)
                        {
                            result.Append(cs.Cells[col, row].FormattedValue + ",");
                        }
                        result.Length--;
                        result.Append(";");
                    }

                }
                //    }
                //}
                return Json(result.ToString());
            }
        }

        private string BuildCommand(string[] measures, string[] dimensions)
        {
            System.Text.StringBuilder result = new System.Text.StringBuilder();
            result.Append("Select {");
            foreach (var measure in measures)
            {
                result.Append($"[Measures].[{measure}], ");
            }
            result.Length--;
            result.Length--;
            result.Append("}");
            result.Append(" ON COLUMNS, NON EMPTY(");
            foreach (var dimension in dimensions)
            {
                result.Append("{");
                result.Append($"{dimension}");
                result.Append("}, ");
            }
            result.Length--;
            result.Length--;
            result.Append(") ON ROWS");
            result.Append($" FROM {CUBE}");

            return result.ToString();

        }

        [Route("api/getByGender")]
        [Route("api/getByAgeBand")]
        [Route("api/getBySegment")]


        /**
        * 2-factor queries
        */
        [Route("api/getByCampaignAndGender")]
        [Route("api/getByCampaignAndAge")]
        [Route("api/getByCampaignAndSegment")]
        [Route("api/getByGenderAndAge")]
        [Route("api/getByGenderAndSegment")]


        /**
        * 3-factor queries
        */

        [Route("api/getByCampaignGenderAndAge")]


        [Route("SSAS/getByCampaign/{measure}")]
        public IActionResult getByCampaign(string measure)
        {
            string connectionString = CNX_STRING;

            List<List<string>> resultData = new List<List<string>>();
            using (AdomdConnection conn = new AdomdConnection(connectionString))
            {
                conn.Open();

                using (AdomdCommand cmd = conn.CreateCommand())
                {
                    // Build the MDX query dynamically based on the provided measure and dimension
                    string mdxQuery = $"SELECT [Measures].[{measure}] ON COLUMNS, NON EMPTY([Campaigns].[Campaign Name].[Campaign Name]) ON ROWS FROM [Comms DW]";

                    cmd.CommandText = mdxQuery;

                    using (AdomdDataReader reader = cmd.ExecuteReader())
                    {
                        // Process the results and store them in the list
                        while (reader.Read())
                        {
                            List<string> row = new List<string>();

                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                row.Add(reader[i].ToString());
                            }

                            resultData.Add(row);
                        }
                    }
                }
            }

            // Return the result data as JSON
            return Json(resultData);
        }


        public IActionResult ExecuteMDXQuery()
            {
                string connectionString = _configuration.GetConnectionString("SSASConnection");

                using (AdomdConnection conn = new AdomdConnection(connectionString))
                {
                    conn.Open();

                    using (AdomdCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "Select {[Measures].[Views]} On columns, ( {[Campaigns].[Campaign Name]}) on rows From [Comms DW]";

                        using (AdomdDataReader reader = cmd.ExecuteReader())
                        {
                            // Process the results
                        }
                    }
                }

                return View();
            }


            public IActionResult TestSSASConnection()
            {
            string connectionString = "Provider=MSOLAP; Data Source=CFEKI_W_P;Initial Catalog=Comms_DW;";
            Console.WriteLine(connectionString);
            List<string> resultData = new List<string>();  // This list will store your result data


            using (AdomdConnection conn = new AdomdConnection(connectionString))
                {
                    conn.Open();

                    using (AdomdCommand cmd = conn.CreateCommand())
                    {
                        // A simple MDX query to retrieve data. Replace this with an actual query based on your cube structure.
                        cmd.CommandText = "Select {[Measures].[Views]} On columns, ( {[Campaigns].[Campaign Name]}) on rows From [Comms DW]";

                        using (AdomdDataReader reader = cmd.ExecuteReader())
                        {
                        while (reader.Read())
                        {
                            resultData.Add(reader[0].ToString());  // Adjust the index based on your actual result structure
                        }
                    }
                    }
                }

                return View(resultData);
            }

        }


    }
