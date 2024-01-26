using Microsoft.AspNetCore.Mvc;
using Microsoft.AnalysisServices.AdomdClient;

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

        private string ExecuteMDXQuery(string[] measures, string[] dimensions, string cube)
        {
            string connectionString = CNX_STRING;
            System.Text.StringBuilder result = new System.Text.StringBuilder();
            using (AdomdConnection conn = new AdomdConnection(connectionString))
            {
                conn.Open();

                using (AdomdCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = BuildCommand(measures, dimensions);
                    Console.WriteLine(cmd.CommandText);

                    CellSet cs = cmd.ExecuteCellSet();

                    foreach (var dim in dimensions)
                    {
                        result.Append(",");
                    }

                    TupleCollection tuplesOnColumns = cs.Axes[0].Set.Tuples;
                    foreach (Microsoft.AnalysisServices.AdomdClient.Tuple column in tuplesOnColumns)
                    {
                        result.Append(column.Members[0].Caption + ",");
                    }
                    result.Length--;
                    result.Append(";");

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

            }
            return result.ToString();
        }


        /**
         * count queries for the homepage
         **/
        [Route("api/count/{measure}")]

        public IActionResult countQuery(string measure)
        {
            string[] measures = { measure };
            string[] dimensions = { "[Campaigns].[Campaign Name]" };
            return Json(ExecuteMDXQuery(measures, dimensions, CUBE));
        }

        /**
        * Simple queries by a single factor 
        */
        [Route("api/getByCampaign")]
        public IActionResult GetByCampaign(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Campaigns].[Campaign Name].[Campaign Name]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }



        [Route("api/getByGender")]
        public IActionResult GetByGender(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Genders].[Gender ID].[Gender ID]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }
        [Route("api/getByAgeBand")]
        public IActionResult GetByAgeBand(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Age Bands].[Age ID].[Age ID]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }
        [Route("api/getBySegment")]
        public IActionResult GetBySegment(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Segments].[Segment].[Segment]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }

        /**
        * 2-factor queries
        */
        [Route("api/getByCampaignAndGender")]
        public IActionResult GetByCampaignAndGender(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Campaigns].[Campaign Name].[Campaign Name]", "[Genders].[Gender ID].[Gender ID]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }
        [Route("api/getByCampaignAndAge")]
        public IActionResult GetByCampaignAndAge(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Campaigns].[Campaign Name].[Campaign Name]", "[Age Bands].[Age ID].[Age ID]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }
        [Route("api/getByCampaignAndSegment")]
        public IActionResult GetByCampaignAndSegment(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Campaigns].[Campaign Name].[Campaign Name]", "[Segments].[Segment].[Segment]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }
        [Route("api/getByGenderAndAge")]
        public IActionResult GetByGenderAndAge(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Genders].[Gender ID].[Gender ID]", "[Age Bands].[Age ID].[Age ID]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }
        [Route("api/getByGenderAndSegment")]
        public IActionResult GetByGenderAndSegmet(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Genders].[Gender ID].[Gender ID]", "[Segments].[Segment].[Segment]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }

        /**
        * 3-factor queries
        */

        [Route("api/getByCampaignGenderAndAge")]
        public IActionResult GetByCampaignGenderAndAge(string measures)
        {
            string[] split_measures = measures.Split(',');
            string[] dimensions = { "[Campaigns].[Campaign Name].[Campaign Name]", "[Genders].[Gender ID].[Gender ID]", "[Age Bands].[Age ID].[Age ID]" };
            return Json(ExecuteMDXQuery(split_measures, dimensions, CUBE));
        }


    }
}
