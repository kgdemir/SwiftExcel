using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;

namespace SwiftExcel
{
    public class ExcelWriter : ExcelWriterCore
    {
        public ExcelWriter(string filePath, Sheet sheet = null, List<int> ColumnWidths = null)
            : base(filePath, sheet, ColumnWidths)
        {
        }
        public void Write(DateTime value, int col, int row)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write($"<c s=\"{(int)ExcelStyles.SHORTDATE}\"");
            if ((value - value.Date).TotalMilliseconds > 0)
            {
                Sheet.Write("><v>");
                Sheet.Write(value.ToOADate().ToString());
            }
            else
            {
                Sheet.Write(" t=\"d\"");
                Sheet.Write("><v>");
                Sheet.Write(value.ToString("yyyy-MM-dd"));
            }
            Sheet.Write("</v></c>");
            Sheet.Write("\n");
            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }
        public void Write(double value, int col, int row)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write($"<c s=\"{(int)ExcelStyles.FINANCIAL_DECIMAL_2}\"");
            Sheet.Write("><v>");
            Sheet.Write(value.ToString().Replace(",", "."));
            Sheet.Write("</v></c>");

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }
        public void Write(long value, int col, int row)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write($"<c s=\"{(int)ExcelStyles.INT}\"");
            Sheet.Write("><v>");
            Sheet.Write(value.ToString());

            Sheet.Write("</v></c>");
            Sheet.Write("\n");
            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }
        public void Write(object o_value, int col, int row)
        {
            if (o_value == null || o_value == DBNull.Value)
            {
                Write("", col, row);
                return;
            }
            if (o_value is DateTime)
            {
                Write((DateTime)o_value, col, row);
                return;
            }
            if (o_value is bool)
            {
                Write(Convert.ToDouble(o_value) > 0 ? 1.0 : 0.0, col, row);
                return;
            }
            if (o_value is byte || o_value is short || o_value is int || o_value is long
                || o_value is ushort || o_value is uint || o_value is ulong)
            {
                Write(Convert.ToInt64(o_value), col, row);
                return;
            }
            if (o_value is decimal || o_value is float || o_value is double)
            {
                Write(Convert.ToDouble(o_value), col, row);
                return;
            }
            Write(Convert.ToString(o_value), col, row);
        }
        public void Write(decimal value, int col, int row)
        {
            Write(Convert.ToDouble(value), col, row);
        }

        public void Write(bool value, int col, int row)
        {
            Write(value ? 1.0 : 0.0, col, row);
        }
        public void Write(byte value, int col, int row)
        {
            Write((double)value, col, row);
        }
        public void Write(short value, int col, int row)
        {
            Write((double)value, col, row);
        }
        public void Write(int value, int col, int row)
        {
            Write((double)value, col, row);
        }


        public void Write(float value, int col, int row)
        {
            Write((double)value, col, row);
        }
        public void Write(string value, int col, int row, DataType dataType = DataType.Text)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write("<c");
            if (dataType == DataType.Text)
            {
                Sheet.Write(" t=\"str\"");
            }
            Sheet.Write("><v>");
            Sheet.Write(EscapeInvalidChars(value));
            Sheet.Write("</v></c>");

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }

        public void WriteFormula(FormulaType type, int col, int row, int sourceCol, int sourceRowStart, int sourceRowEnd)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write("<c><f>");
            Sheet.Write($"{type.ToString().ToUpper()}");
            Sheet.Write("(");
            Sheet.Write($"{GetFullCellName(sourceCol, sourceRowStart)}");
            Sheet.Write(":");
            Sheet.Write($"{GetFullCellName(sourceCol, sourceRowEnd)}");
            Sheet.Write(")</f><v></v></c>");

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }

        private static string GetFullCellName(int col, int row)
        {
            return $"{GetCellName(col)}{row}";
        }

        private static string GetCellName(int col)
        {
            var dividend = col;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = (char)(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private static string EscapeInvalidChars(string value)
        {
            value = SecurityElement.Escape(value);
            if (!Configuration.UseEnchancedXmlEscaping)
            {
                return value;
            }

            if (string.IsNullOrEmpty(value) || value.All(XmlConvert.IsXmlChar))
            {
                return value;
            }

            var result = new StringBuilder();
            foreach (var character in value)
            {
                if (XmlConvert.IsXmlChar(character))
                {
                    result.Append(character);
                }
                else
                {
                    result.Append($"_x{(int)character:x4}_");
                }
            }

            return result.ToString();
        }
    }
}