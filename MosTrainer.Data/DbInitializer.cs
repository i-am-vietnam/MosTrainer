using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SQLite;

namespace MosTrainer.Data
{
    public class DbInitializer
    {
        public void EnsureCreated()
        {
            string cs = ConfigurationManager.ConnectionStrings["MosDb"].ConnectionString;

            using (var conn = new SQLiteConnection(cs))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS TaskResults (
  Id INTEGER PRIMARY KEY AUTOINCREMENT,
  Username TEXT NOT NULL,
  ProjectId TEXT NOT NULL,
  TaskId TEXT NOT NULL,
  IsPass INTEGER NOT NULL,
  CheckedAt TEXT NOT NULL
);";
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}

