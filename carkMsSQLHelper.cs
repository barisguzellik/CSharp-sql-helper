using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace carkSQLHelper
{
    public class carkMsSQLHelper
    {
        //DEFAULT VARIABLES
        SqlConnection sqlConnectionSource = new SqlConnection(ConfigurationManager.ConnectionStrings["carkMsSQLHelper"].ConnectionString);

        public DataTable getDataTable(string sqlQuery)
        {
            SqlDataAdapter sqlDataAdapterItem = new SqlDataAdapter(sqlQuery, sqlConnectionSource);
            DataTable dataTableItem = new DataTable();
            sqlDataAdapterItem.Fill(dataTableItem);
            return dataTableItem;
        }

        public string getTableColumnData(string sqlQuery, string columnName)
        {
            SqlDataAdapter sqlDataAdapterItem = new SqlDataAdapter(sqlQuery, sqlConnectionSource);
            DataTable dataTableItem = new DataTable();
            sqlDataAdapterItem.Fill(dataTableItem);
            if (dataTableItem.Rows.Count > 0)
            { return dataTableItem.Rows[0][columnName].ToString(); }
            else { return ""; }
        }

        public string getTableRowCount(string sqlQuery)
        {
            SqlDataAdapter sqlDataAdapterItem = new SqlDataAdapter(sqlQuery, sqlConnectionSource);
            DataTable dataTableItem = new DataTable();
            sqlDataAdapterItem.Fill(dataTableItem);
            if (dataTableItem.Rows.Count > 0)
            { return dataTableItem.Rows[0][0].ToString(); }
            else { return "NaN"; }
        }

        public string getTableRowCountForCountQuery(string tableName, string countColumn, string whereQuery)
        {
            string sqlQuery = "SELECT COUNT("+countColumn+") FROM "+tableName+" WHERE "+whereQuery+"";
            SqlDataAdapter sqlDataAdapterItem = new SqlDataAdapter(sqlQuery, sqlConnectionSource);
            DataTable dataTableItem = new DataTable();
            sqlDataAdapterItem.Fill(dataTableItem);
            if (dataTableItem.Rows.Count > 0)
            { return dataTableItem.Rows[0][0].ToString(); }
            else { return "NaN"; }
        }

        public DataTable getDataTableWithParameters(string sqlQuery,string[] parameters,string[] values)
        {
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery,sqlConnectionSource);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommandItem.Parameters.AddWithValue(parameters[i], values[i]);
            }
            sqlCommandItem.Connection.Open();
            SqlDataReader sqlDataReaderItem = sqlCommandItem.ExecuteReader();
            DataTable dataTableItem = new DataTable();
            dataTableItem.Load(sqlDataReaderItem);
            sqlConnectionSource.Close();
            return dataTableItem;
        }

        public DataTable getDataTableForLikeQueryWithParameters(string sqlQuery, string[] parameters, string[] values, string likeColumn, string likeValue)
        {
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommandItem.Parameters.AddWithValue(parameters[i], values[i]);
            }
            sqlCommandItem.Parameters.AddWithValue("@" + likeColumn + "", "%" + likeValue + "%");
            sqlCommandItem.Connection.Open();
            SqlDataReader sqlDataReaderItem = sqlCommandItem.ExecuteReader();
            DataTable dataTableItem = new DataTable();
            dataTableItem.Load(sqlDataReaderItem);
            sqlConnectionSource.Close();
            return dataTableItem;
        }

        public void setSqlCommandWithParameters(string sqlQuery, string[] parameters, string[] values)
        {
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommandItem.Parameters.AddWithValue(parameters[i], values[i]);
            }
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }

        public void setSqlCommand(string sqlQuery)
        {
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }

        /*public void setSqlCommandForInsert(string tableName,string[] columns,string[] values)
        {
            string stringColumns = returnArrayToStringValues(columns);
            string stringValues = returnArrayToStringValues(values);
            string sqlQuery="INSERT INTO "+tableName+"("+stringColumns+") VALUES("+stringValues+")";
            SqlCommand sqlCommandItem=new SqlCommand(sqlQuery,sqlConnectionSource);
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }*/

        public void setSqlCommandForInsertWithParameters(string tableName, string[] columns, string[] values)
        {
            string stringColumns = returnArrayToStringValues(columns);
            string stringValues = returnArrayToStringValues(values);

            string[] parameters = returnColumnsToParameters(columns);
            string stringParameters = returnArrayToStringValues(parameters);

            string sqlQuery = "INSERT INTO " + tableName + "(" + stringColumns + ") VALUES(" + stringParameters + ")";
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommandItem.Parameters.AddWithValue(parameters[i], values[i]);
            }
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }

        /*public void setSqlCommandForUpdate(string tableName, string[] columns, string[] values,string whereColumn, string whereValue)
        {
            string columnValueCompare = "";
            for (int i = 0; i < columns.Length; i++)
            {
                columnValueCompare += columns[i] + "=" + values[i] + ",";
            }
            columnValueCompare = columnValueCompare.Remove(columnValueCompare.Length - 1, 1);

            string sqlQuery = "UPDATE "+tableName+" SET "+columnValueCompare+" WHERE "+whereColumn+"="+whereValue+"";
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }*/

        public void setSqlCommandForUpdateWithParameters(string tableName, string[] columns, string[] values, string whereColumn, string whereValue)
        {
            string[] parameters = returnColumnsToParameters(columns);
            string columnParameterCompare = "";
            for (int i = 0; i < columns.Length; i++)
            {
                columnParameterCompare += columns[i] + "=" + parameters[i] + ",";
            }
            columnParameterCompare = columnParameterCompare.Remove(columnParameterCompare.Length - 1, 1);

            string sqlQuery = "UPDATE " + tableName + " SET " + columnParameterCompare + " WHERE " + whereColumn+"=@VALUE";
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommandItem.Parameters.AddWithValue(parameters[i], values[i]);
            }
            sqlCommandItem.Parameters.AddWithValue("@VALUE", whereValue);
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }

        public void setSqlCommandForDeleteWithParameters(string tableName, string whereColumn, string whereValue)
        {
            string sqlQuery = "DELETE FROM "+tableName+" WHERE "+whereColumn+"=@VALUE";
            SqlCommand sqlCommandItem = new SqlCommand(sqlQuery, sqlConnectionSource);
            sqlCommandItem.Parameters.AddWithValue("@VALUE", whereValue);
            sqlCommandItem.Connection.Open();
            sqlCommandItem.ExecuteNonQuery();
            sqlConnectionSource.Close();
        }

        public string[] getColumnsNames(string tableName)
        {
            List<string> listacolumnas = new List<string>();
            SqlCommand sqlCommandItem = new SqlCommand("select c.name from sys.columns c inner join sys.tables t on t.object_id = c.object_id and t.name = '" + tableName + "' and t.type = 'U'", sqlConnectionSource);
            sqlCommandItem.Connection.Open();
            var reader = sqlCommandItem.ExecuteReader();
            while (reader.Read())
            {
                listacolumnas.Add(reader.GetString(0));
            }
            sqlConnectionSource.Close();
            return listacolumnas.ToArray();
        }

        public string[] returnColumnsToParameters(string[] columns)
        {
            string[] parameters = new string[columns.Length];
            for (int i = 0; i < columns.Length; i++)
            {
                parameters[i] = "@" + columns[i];
            }
            return parameters;
        }

        public string returnArrayToStringValues(string[] array)
        {
            string stringValues = "";
            foreach (var item in array)
            {
                stringValues += item.ToString() + ",";
            }
            stringValues = stringValues.Remove(stringValues.Length - 1, 1);
            return stringValues;
        }

        public string getRandomFileName(string fileExtention)
        {
            string day = DateTime.Now.Day.ToString();
            string month = DateTime.Now.Month.ToString();
            string year = DateTime.Now.Year.ToString();
            string hour = DateTime.Now.Hour.ToString();
            string minutes = DateTime.Now.Minute.ToString();
            string second = DateTime.Now.Second.ToString();
            string milisecond = DateTime.Now.Millisecond.ToString();
            return day + month + year + hour + minutes + second + milisecond + fileExtention;
        }



        public string getLinkCreator(string text)
        {
            string strReturn = text.Trim();
            strReturn = strReturn.ToLower();
            strReturn = strReturn.Replace("ğ", "g");
            strReturn = strReturn.Replace("Ğ", "G");
            strReturn = strReturn.Replace("ü", "u");
            strReturn = strReturn.Replace("Ü", "U");
            strReturn = strReturn.Replace("ş", "s");
            strReturn = strReturn.Replace("Ş", "S");
            strReturn = strReturn.Replace("ı", "i");
            strReturn = strReturn.Replace("İ", "I");
            strReturn = strReturn.Replace("ö", "o");
            strReturn = strReturn.Replace("Ö", "O");
            strReturn = strReturn.Replace("ç", "c");
            strReturn = strReturn.Replace("Ç", "C");
            strReturn = strReturn.Replace("-", "+");
            strReturn = strReturn.Replace(" ", "+");
            strReturn = strReturn.Trim();
            strReturn = new System.Text.RegularExpressions.Regex("[^a-zA-Z0-9+]").Replace(strReturn, "");
            strReturn = strReturn.Trim();
            strReturn = strReturn.Replace("+", "-");
            return strReturn;
        }

        public void setResizeImage(int size, string filePath, string saveFilePath)
        {
            //variables for image dimension/scale
            double newHeight = 0;
            double newWidth = 0;
            double scale = 0;

            //create new image object
            Bitmap curImage = new Bitmap(filePath);

            //Determine image scaling
            if (curImage.Height > curImage.Width)
            {
                scale = Convert.ToSingle(size) / curImage.Height;
            }
            else
            {
                scale = Convert.ToSingle(size) / curImage.Width;
            }
            if (scale < 0 || scale > 1) { scale = 1; }

            //New image dimension
            newHeight = Math.Floor(Convert.ToSingle(curImage.Height) * scale);
            newWidth = Math.Floor(Convert.ToSingle(curImage.Width) * scale);

            //Create new object image
            Bitmap newImage = new Bitmap(curImage, Convert.ToInt32(newWidth), Convert.ToInt32(newHeight));
            Graphics imgDest = Graphics.FromImage(newImage);
            imgDest.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            imgDest.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            imgDest.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            imgDest.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            ImageCodecInfo[] info = ImageCodecInfo.GetImageEncoders();
            EncoderParameters param = new EncoderParameters(1);
            param.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L);

            //Draw the object image
            imgDest.DrawImage(curImage, 0, 0, newImage.Width, newImage.Height);

            //Save image file
            newImage.Save(saveFilePath, info[1], param);

            //Dispose the image objects
            curImage.Dispose();
            newImage.Dispose();
            imgDest.Dispose();
        }

    }
    
}
