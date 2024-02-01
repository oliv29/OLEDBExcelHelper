namespace OLEDBExcelHelper.Contracts
{
    public interface IExcelHelper : IDisposable
    {
        /// <summary>
        /// Inserts a data row in the specified worksheeet
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <param name="dataRowValue">dataRowValue</param>
        /// <returns>Row(s) affected (TODO: update to accept collection)</returns>
        int InsertData(string worksheetName, string dataRowValue);

        /// <summary>
        /// This function basically changes the column type (e.g. from text to memo, so that it accepts more than 255 chars (for Excel 8))
        /// </summary>
        /// <param name="worksheetName">worksheetName</param>
        /// <param name="memoColumns">memoColumns</param>
        void PrepareWorksheet(string worksheetName, IEnumerable<string> memoColumns);
    }
}
