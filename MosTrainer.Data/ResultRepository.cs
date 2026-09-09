using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Configuration;
using System.Data.SQLite;
using System;
using System.Configuration;
using System.Data.SQLite;

namespace MosTrainer.Data
{
    public class ResultRepository
    {
        private string ConnStr
        {
            get { return ConfigurationManager.ConnectionStrings["MosDb"].ConnectionString; }
        }

        public void Save(string username, string projectId, string taskId, bool isPass)
        {
            using (var conn = new SQLiteConnection(ConnStr))
            {
                conn.Execute(@"
INSERT INTO TaskResults(Username, ProjectId, TaskId, IsPass, CheckedAt)
VALUES (@u, @p, @t, @pass, @time)",
                new
                {
                    u = username,
                    p = projectId,
                    t = taskId,
                    pass = isPass ? 1 : 0,
                    time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                });
            }
        }
    }
}
