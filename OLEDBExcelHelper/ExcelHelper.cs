using System.Data;
using System.Data.OleDb;

namespace Benchmarking
{
    internal class ExcelHelper
    {
        private readonly string             _connectionString;
        private readonly OleDbConnection    _connection;

        private const string COLUMN_SCHEMA_NAME = "ColumnName";

        private Dictionary<string, string> _currentWorksheetColumns = new Dictionary<string, string>();

        private Dictionary<string, string> _defaultWorksheetColumns = new Dictionary<string, string>() { 
            { "Auftragsnummer", "TEXT" },
            { "Auftragsdatum", "TEXT" },
            { "Leistungsdatum", "TEXT" },
            { "Auftragsstatus", "TEXT" },
            { "Stornierte Auftragsnummer", "TEXT" },
            { "Händlernummer", "TEXT" },
            { "Anfallstelle", "MEMO" },
            { "Systempartner", "TEXT" },
            { "Auftragszeile", "TEXT" },
            { "Leistungsnummer", "TEXT" },
            { "Leistungsbeschreibung", "TEXT" },
            { "Menge", "TEXT" },
            { "Bemerkungen", "MEMO" },
        };

        public ExcelHelper(string connectionString)
        {
            _connectionString = connectionString;                
        }

        public ExcelHelper(OleDbConnection connection)
        {
            _connection = connection;                
        }

        public int InsertData(string worksheetName, string dataRowValue)
        {
            using (OleDbCommand dbCommand = _connection.CreateCommand())
            {
                var sqlCommandQuery = string.Format("INSERT INTO [{0}] VALUES ({1})", worksheetName, dataRowValue);

                return ExecuteQuery(dbCommand, sqlCommandQuery);
            }
        }

        public void PrepareWorksheet(string worksheetName, IEnumerable<string> memoColumns)        
        {    
            worksheetName = FormatWorksheetName(worksheetName);

            if (WorksheetExists(worksheetName)) 
            {
                var existingHeaderColumns   = GetColumnsNames(worksheetName);
                var newColumns              = CreateNewColumns(existingHeaderColumns, memoColumns);

                DropWorksheet(worksheetName);

                CreateNewWorksheet(worksheetName, newColumns);
            }
            else
            {
                CreateNewWorksheet(worksheetName, _currentWorksheetColumns);
            }
        }

        private void CreateNewWorksheet(string worksheetName, Dictionary<string, string> columnsToAdd)
        {
            using (OleDbCommand dbCommand = _connection.CreateCommand())
            {
                var commandQuery = string.Format("CREATE TABLE [{0}] (", worksheetName);

                foreach (var column in SelectValidColumnsSet(columnsToAdd))
                {
                    commandQuery += string.Format("[{0}] {1},", column.Key, column.Value);
                }

                commandQuery = commandQuery.Substring(0, commandQuery.Length - 1) + ")";

                ExecuteQuery(dbCommand, commandQuery);
            }
        }

        private Dictionary<string, string> SelectValidColumnsSet(Dictionary<string, string> computedColumns)
        {
            return computedColumns.Count == 0
                ? (_currentWorksheetColumns.Count == 0) ? _defaultWorksheetColumns : _currentWorksheetColumns
                : computedColumns;
        }

        private Dictionary<string, string> CreateNewColumns(IEnumerable<string> worksheetColumns, IEnumerable<string> memoColumns)
        {
            var dctColumnsToAdd = new Dictionary<string, string>(); 

            foreach (var column in worksheetColumns)
            {
                dctColumnsToAdd[column] = memoColumns.Contains(column) ? "MEMO" : "TEXT";
            }

            if (dctColumnsToAdd.Count > 0)
                _currentWorksheetColumns = dctColumnsToAdd;

            return dctColumnsToAdd;
        }

        private void DropWorksheetColumns(string worksheetName, IEnumerable<string> columnsToDrop)
        {
            using (OleDbCommand dbCommand = _connection.CreateCommand())
            {
                var commandText = string.Format("ALTER TABLE [{0}] DROP COLUMN ", worksheetName);

                foreach (var column in columnsToDrop)
                {
                    commandText += string.Format("[{0}],", column);
                }

                commandText = commandText.Substring(0, commandText.Length - 1);

                ExecuteQuery(dbCommand, commandText);
            }
        }

        private void DropWorksheet(string worksheetName)
        {
            using (OleDbCommand dbCommand = _connection.CreateCommand())
            {
                var commandText = string.Format("DROP TABLE [{0}]", worksheetName);

                ExecuteQuery(dbCommand, commandText);
            }
        }

        private void AddWorksheetColumns(string worksheetName, Dictionary<string, string> columnsToAdd)
        {
            using (OleDbCommand dbCommand = _connection.CreateCommand())
            {
                var commandText = string.Format("ALTER TABLE [{0}$] ADD COLUMN ", worksheetName);

                foreach (var column in columnsToAdd)
                {
                    commandText += string.Format("[{0}] {1},", column.Key, column.Value);
                }

                commandText = commandText.Substring(0, commandText.Length - 1);

                ExecuteQuery(dbCommand, commandText);
            }
        }

        private bool WorksheetExists(string worksheetName)
        {
            var existingWorksheets = GetWorksheetsNames(_connection);

            return existingWorksheets.Any(x => x.Equals(worksheetName));
        }

        private IEnumerable<string> GetWorksheetsNames(OleDbConnection connection)
        {
            var existingWorksheets = new List<string>();

            try
            {
                DataTable dtSheetsNames = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" }) ?? throw new InvalidOperationException("Could not retrieve worksheets!");

                foreach (DataRow row in dtSheetsNames.Rows)
                {
                    existingWorksheets.Add(row["TABLE_NAME"].ToString());
                }

                return existingWorksheets;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("An error occcured on worksheets retreival [{0}]", ex.Message));
            }
        }

        private IEnumerable<string> GetColumnsNames(string worksheetName)
        {
            var existingColumns = new List<string>();

            try
            {
                using (var command = new OleDbCommand(string.Format("SELECT * FROM [{0}]", worksheetName), _connection))
                {
                    using (var reader = command.ExecuteReader(CommandBehavior.SchemaOnly))
                    {
                        var tableSchema = reader.GetSchemaTable();
                        var nameCol     = tableSchema.Columns[COLUMN_SCHEMA_NAME];

                        foreach (DataRow row in tableSchema.Rows)
                        {
                            existingColumns.Add(row[nameCol].ToString());
                        }
                    }
                }

                if (existingColumns.Count != _defaultWorksheetColumns.Count) 
                { 
                    return Enumerable.Empty<string>();
                }

                return existingColumns;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("An error occured on columns fetching [{0}]", ex.Message));
            }
        }

        private int ExecuteQuery(OleDbCommand command, string queryToExecute)
        {
            if (string.IsNullOrWhiteSpace(queryToExecute))
                throw new ArgumentNullException("queryToExecute");

            command.CommandText = queryToExecute;
            command.CommandType = CommandType.Text;

            return command.ExecuteNonQuery();
        }

        private string FormatWorksheetName(string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(worksheetName))
                throw new ArgumentNullException("worksheetName");

            if (worksheetName.Substring(worksheetName.Length - 1) != "$")
                worksheetName += "$";

            return worksheetName;
        }
    }
}
