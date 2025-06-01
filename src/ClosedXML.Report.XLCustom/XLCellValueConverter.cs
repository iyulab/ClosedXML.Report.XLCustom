namespace ClosedXML.Report.XLCustom;

internal static class XLCellValueConverter
{
    public static XLCellValue FromObject(object obj, IFormatProvider provider = null)
    {
        if (obj != null)
        {
            if (!(obj is XLCellValue result))
            {
                if (!(obj is Blank blank))
                {
                    if (!(obj is bool flag))
                    {
                        if (!(obj is string text))
                        {
                            if (!(obj is XLError xLError))
                            {
                                if (!(obj is DateTime dateTime))
                                {
                                    if (!(obj is TimeSpan timeSpan))
                                    {
                                        if (!(obj is sbyte b))
                                        {
                                            if (!(obj is byte b2))
                                            {
                                                if (!(obj is short num))
                                                {
                                                    if (!(obj is ushort num2))
                                                    {
                                                        if (!(obj is int num3))
                                                        {
                                                            if (!(obj is uint num4))
                                                            {
                                                                if (!(obj is long num5))
                                                                {
                                                                    if (!(obj is ulong num6))
                                                                    {
                                                                        if (!(obj is float num7))
                                                                        {
                                                                            if (!(obj is double num8))
                                                                            {
                                                                                if (obj is decimal num9)
                                                                                {
                                                                                    return num9;
                                                                                }

                                                                                return Convert.ToString(obj, provider);
                                                                            }

                                                                            return num8;
                                                                        }

                                                                        return num7;
                                                                    }

                                                                    return num6;
                                                                }

                                                                return num5;
                                                            }

                                                            return num4;
                                                        }

                                                        return num3;
                                                    }

                                                    return num2;
                                                }

                                                return num;
                                            }

                                            return b2;
                                        }

                                        return b;
                                    }

                                    return timeSpan;
                                }

                                return dateTime;
                            }

                            return xLError;
                        }

                        return text;
                    }

                    return flag;
                }

                return blank;
            }

            return result;
        }

        return Blank.Value;
    }
}