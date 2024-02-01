using OLEDBExcelHelper.Contracts;
using System.Data;
using System.Data.OleDb;

namespace Benchmarking
{
    internal class ExcelHelper : IExcelHelper
    {
        #region Fields

        private readonly string _connectionString;
        private OleDbConnection _connection;

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
            { "Bemerkungen", "MEMO" }
        };

        #endregion Fields

        #region Instance

        /// <summary>
        /// Public constructor
        /// </summary>
        /// <param name="connectionString">connectionString</param>
        public ExcelHelper(string connectionString)
        {
            _connectionString = connectionString ?? throw new ArgumentNullException("Provided connection string is null or empty!");                
        }

        #endregion Instance

        #region Public section

        /// <summary>
        /// Inserts a data row in the specified worksheeet
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <param name="dataRowValue">dataRowValue</param>
        /// <returns>Row(s) affected (TODO: update to accept collection)</returns>
        public int InsertData(string worksheetName, string dataRowValue)
        {
            using (OleDbCommand dbCommand = Connection.CreateCommand())
            {
                var sqlCommandQuery = string.Format("INSERT INTO [{0}] VALUES ({1})", worksheetName, dataRowValue);

                return ExecuteQuery(dbCommand, sqlCommandQuery);
            }
        }

        /// <summary>
        /// This function basically changes the column type (e.g. from text to memo, so that it accepts more than 255 chars (for Excel 8))
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <param name="memoColumns">memoColumns</param>
        public void PrepareWorksheet(string worksheetName, IEnumerable<string> memoColumns)        
        {
            var newWorksheetName = worksheetName;

            worksheetName = FormatWorksheetName(worksheetName);

            if (WorksheetExists(worksheetName)) 
            {
                UpdateWorksheet(worksheetName, memoColumns);
            }
            else
            {
                CreateWorksheet(newWorksheetName, _currentWorksheetColumns);
            }
        }

        public void Dispose()
        {
            if (Connection != null)
            {
                if (Connection.State != ConnectionState.Closed)
                    Connection.Close(); //will automatically call Dispose()

                _connection = null;
            }
        }

        #endregion Public section

        #region Private section

        private void UpdateWorksheet(string worksheetName, IEnumerable<string> memoColumns)
        {
            var existingHeaderColumns = GetColumnsNames(worksheetName);

            if (existingHeaderColumns.Any())
                ClearWorksheet(worksheetName);

            var newColumns = CreateNewColumns(existingHeaderColumns, memoColumns);

            CreateWorksheet(worksheetName, newColumns);
        }

        private void CreateWorksheet(string worksheetName, Dictionary<string, string> columnsToAdd)
        {
            using (OleDbCommand dbCommand = Connection.CreateCommand())
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

        /// <summary>
        /// Clears worksheet's data (including header columns)
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        private void ClearWorksheet(string worksheetName)
        {
            using (OleDbCommand dbCommand = Connection.CreateCommand())
            {
                var commandText = string.Format("DELETE FROM [{0}]", worksheetName);

                ExecuteQuery(dbCommand, commandText);
            }
        }

        /// <summary>
        /// Checks if a worksheet specified by name, exists in the current workbook
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <returns>True if worksheet exists, false otherwise</returns>
        private bool WorksheetExists(string worksheetName)
        {
            var existingWorksheets = GetWorksheetsNames();

            return existingWorksheets.Any(x => x.Equals(worksheetName));
        }

        private IEnumerable<string> GetWorksheetsNames()
        {
            var existingWorksheets = new List<string>();

            try
            {
                DataTable dtSheetsNames = Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" }) ?? throw new InvalidOperationException("Could not retrieve worksheets!");

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
                using (var command = new OleDbCommand(string.Format("SELECT * FROM [{0}]", worksheetName), Connection))
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

                return (existingColumns.Count != _defaultWorksheetColumns.Count) ? Enumerable.Empty<string>() : existingColumns;
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

        private OleDbConnection Connection
        {
            get
            {
                if (_connection == null)
                    _connection = new OleDbConnection(_connectionString);

                if (_connection.State != ConnectionState.Open)
                    _connection.Open();

                return _connection;
            }
        }

        #region Not used

        /// <summary>
        /// Drops worksheet's columns
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <param name="columnsToDrop">columnsToDrop</param>
        private void DropWorksheetColumns(string worksheetName, IEnumerable<string> columnsToDrop)
        {
            using (OleDbCommand dbCommand = Connection.CreateCommand())
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

        /// <summary>
        /// Adds columns to an existing worksheet
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <param name="columnsToAdd">columnsToAdd</param>
        private void AddWorksheetColumns(string worksheetName, Dictionary<string, string> columnsToAdd)
        {
            using (OleDbCommand dbCommand = Connection.CreateCommand())
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

        #endregion Not uses

        #endregion Private section
    }
}
