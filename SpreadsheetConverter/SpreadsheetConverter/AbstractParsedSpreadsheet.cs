///Author:Ben Millane
///http://stackoverflow.com/users/2539260/
///https://github.com/benmillane/Portfolio
///Date: 06.03.2014
///Usage: Get desired spreadsheet to import
///Create concrete implementation of IFileStorage - populate the StorageLocation property with a file path for your sheet. (xls only)
///Create concrete implementation of IParsedRow for your sheet, this will be used to represent a row. It should contain a list of properties
///which names are in lower case an match exactly what is in your spreadsheet.
///Create a class which implements AbstractParsedSpreadsheet and calls its default constructor. Override any functionality you
///need to and add in any additional functionality as required.
///Create a map which needs to be a dictionary of string to int. Each string should be lower case and match both a column name and 
///property in your IParsedRow implementation. The int needs to be the column index of said column name which is used when assigining
///cell values to properties via reflection.
///
///Thanks to the NPOI team for creating a very nice wrapper around the horrible .NET spreadsheet classes.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using SpreadsheetConverter.Interfaces;
using SpreadsheetConverter.Files;

namespace SpreadsheetConverter.AbstractClasses
{
    /// <summary>
    /// Abstract, generic class that only accepts a type parameter which implements IParsedRow for use in its internal row collection.
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    public abstract class AbstractParsedSpreadsheet<TEntity> : IParsedSpreadsheet<TEntity>, IEnumerable<TEntity> where TEntity : IParsedRow, new()
    {
        public int Columns { get; set; }
        public int Pages { get; set; }
        public int InterfacePropertyCount { get; set; }
        public Dictionary<string, int> Map { get; set; }
        public List<string> ColumnHeaders { get; set; }
        public List<TEntity> RowList { get; set; }

        /// <summary>
        /// Single parametised constructor which will perform all work needed to make an object from a spreadsheet.
        /// Used on the front end this might have a dropdown list of all "maps" or spreadsheet types you can upload.
        /// and then the act of uploading will create a storage item. If the spreadsheet in the storage item doesnt match
        /// up to the map then an error will be thrown.
        /// </summary>
        /// <param name="validationMap">this should hold the names of the columns in the spreadsheet (and also the TEntities properties) and the column index in which the column name resides in</param>
        /// <param name="storage">this object has a property which refers to a file location and is used by NPOI to load up a spreadsheet for checking and parsing.</param>
        public AbstractParsedSpreadsheet(Dictionary<string,int> validationMap, IFileStorage storage)
        {
            //Get a count of all properties that the IParsedRow interface contains (used to subtract from expected amount when mapping).
            this.InterfacePropertyCount = typeof(TEntity).GetInterfaces()[0].GetProperties().Count();
            this.Map = validationMap;

            //Check validation map against properties of TEntity
            if (!this.CheckMapMatchesRowType())
            {
                throw new InvalidDataException("Invalid Map/Type parameter used");
            }
            
            //Obtain column headers from spreadsheet
            this.ColumnHeaders = ObtainColumnHeaders(storage);

            //Check validationMap against column headers
            if (!CheckHeadersAgainstMap())
            {
                throw new InvalidDataException("Invalid Spreadsheet/Map used");
            }

            //Parse spreadsheet into RowList if all of the above pass.
            this.RowList = ParseSheet(storage);
        }

        /// <summary>
        /// This method takes in an IFileStorage implementation and uses it to locate and open a spreadsheet.
        /// It then reads from the spreadsheet, calling another function to create objects of type TEntity
        /// and adds them into a list which belongs to this class.
        /// </summary>
        /// <param name="storage"></param>
        /// <returns></returns>
        public virtual List<TEntity> ParseSheet(IFileStorage storage)
        {
            List<TEntity> ListOfRows = new List<TEntity>();

            HSSFWorkbook hssfbook;

            using (FileStream file = new FileStream(storage.StorageLocation, FileMode.Open, FileAccess.Read))
            {
                hssfbook = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfbook.GetSheetAt(0);

            foreach (IRow row in sheet)
            {
                if (row.RowNum == 0)
                {
                    continue;
                }
                else
                {

                    ListOfRows.Add(CreateEntityFromRow(row));
                }
            }

            return ListOfRows;
        }

        /// <summary>
        /// Bit of a complicated one - Accepts an IRow implementing object (those used by the NPOI spreadsheet classes)
        /// looks up the column index of each cell in a row and maps it using the local Map variable (dictionary of string to int)
        /// to a string value. This value can then be used to dynamically obtain a property name from TEntity using .NET Reflection.
        /// The value of the current cell is then set to that property on TEntity before being continuing to the next cell.
        /// After the entire object is populated it returns it.
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public virtual TEntity CreateEntityFromRow(IRow row)
        {
            TEntity retVal = new TEntity();
            Type entity = typeof(TEntity);

            foreach (ICell c in row)
            {
                //Looks up the column index of the current cell and Maps it to the corresponding value in the Map dictionary to 
                //obtain the correct property name in TEntity that this value needs to be set for.
                string columnName = this.Map.Where(d => d.Value == c.ColumnIndex).Select(e => e.Key).First();

                switch (c.CellType)
                {
                    case CellType.STRING:
                        retVal.GetType().GetProperty(columnName).SetValue(retVal, c.StringCellValue.ToString(), null);
                        break;
                    case CellType.NUMERIC:
                        retVal.GetType().GetProperty(columnName).SetValue(retVal, c.NumericCellValue, null);
                        break;
                    case CellType.BOOLEAN:
                        retVal.GetType().GetProperty(columnName).SetValue(retVal, c.BooleanCellValue, null);
                        break;
                    case CellType.BLANK:
                    case CellType.ERROR:
                    case CellType.FORMULA:
                    case CellType.Unknown:
                    default:
                        break;
                }
                //Hardcoded - will refactor this later to use a loop and place in own method.
                retVal.GetType().GetProperty("RowNumber").SetValue(retVal, row.RowNum, null);
            }

            return retVal;

        }
        /// <summary>
        /// Looks up the generic parameter for this class, instatiates it and checks that its properties match the map.
        /// It then checks to ensure that the map contains the correct number of entries for the number of properties on
        /// the generic type.
        /// </summary>
        /// <returns></returns>
        public virtual bool CheckMapMatchesRowType()
        {
            Type entity = typeof(TEntity);

            //list of all properties (including inherited properties) for TEntity.
            var properties = entity.GetProperties().ToList();

            //Empty list to hold all properties exclusive to TEntity (not including inherited properties).
            List<PropertyInfo> propInfo = new List<PropertyInfo>();

            //Loops through all of the properties in TEntity. Cross checks the name of each property against the name of
            //each property in IParsedRow (which is its first interface). If the name of the property in TEntity is also
            //found as a property name in IParsedRow then it will NOT add the property to the propInfo list as it is not
            //a property which needs to be mapped.
            foreach (var i in properties)
            {
                if (typeof(TEntity).GetInterfaces()[0].GetProperties().ToList().Select(c => c.Name).Contains(i.Name))
                {
                    continue;
                }
                else
                {
                    propInfo.Add(i);
                }
            }

            if (propInfo.Count() != Map.Count)
            {
                return false;
            }

            //Check each property name from TEntity against the values in the map
            //If any don't match up then return false.
            foreach (var i in propInfo)
            {
                if (!Map.Keys.Contains(i.Name.ToLower())){
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Gets the top row of any spreadsheet (which is normally where the headers are)
        /// </summary>
        /// <param name="storage"></param>
        /// <returns></returns>
        public virtual List<string> ObtainColumnHeaders(IFileStorage storage)
        {
            HSSFWorkbook hssfbook;
            List<string> ColumnHeaders = new List<string>();

            using (FileStream file = new FileStream(storage.StorageLocation, FileMode.Open, FileAccess.Read))
            {
                hssfbook = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfbook.GetSheetAt(0);
            IRow row = sheet.GetRow(0);

            foreach (ICell c in row)
            {
                switch (c.CellType)
                {
                    case CellType.STRING:
                        ColumnHeaders.Add(c.StringCellValue.ToString().Replace(" ", string.Empty));
                        break;
                    case CellType.NUMERIC:
                    case CellType.BOOLEAN:
                    case CellType.BLANK:
                    case CellType.ERROR:
                    case CellType.FORMULA:
                    case CellType.Unknown:
                    default:
                        break;
                }

            }

            return ColumnHeaders;
        }

        /// <summary>
        /// Checks that the headers obtained from the spreadsheet passed in are valid against the map that has been passed in
        /// also checks that the count of both of them matches.
        /// 
        /// </summary>
        /// <returns></returns>
        public virtual bool CheckHeadersAgainstMap(){
            if (ColumnHeaders.Count != this.Map.Values.Count)
            {
                return false;
            }
            foreach (string i in this.ColumnHeaders)
            {
                if (!this.Map.Keys.Contains(i.ToLower()))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Make the RowList propert of the class it's enumerable.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<TEntity> GetEnumerator()
        {
            foreach (TEntity t in this.RowList)
            {
                if (t == null)
                {
                    break;
                }

                yield return t;
            }
        }

        [ExcludeFromCodeCoverage]
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
