using ClosedXML.Report.Options;
using System.Globalization;

namespace ClosedXML.Report.XLCustom.Tags;

/// <summary>
/// Tag that applies standard .NET format strings to values
/// </summary>
public class FormatTag : OptionTag
{
    public override void Execute(ProcessingContext context)
    {
        var xlCell = Cell.GetXlCell(context.Range);

        try
        {
            var variableName = GetParameter("name");
            var formatString = GetParameter("format");

            Log.Debug($"FormatTag - name: {variableName}, format: {formatString}");

            if (string.IsNullOrEmpty(variableName))
            {
                Log.Debug("FormatTag - Missing variable name parameter");
                xlCell.Value = "Error: Missing variable name";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            if (string.IsNullOrEmpty(formatString))
            {
                Log.Debug("FormatTag - Missing format parameter");
                xlCell.Value = "Error: Missing format string";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // 변수 평가 - Parameter 객체를 생성하여 전달
            object? value = null;
            try
            {
                // context.Parameters 대신 Parameter 객체를 직접 생성하여 전달
                value = context.Evaluator.Evaluate(variableName, new Parameter("item", context.Value));
                Log.Debug($"Evaluated variable {variableName} = {value ?? "null"}");
            }
            catch (Exception ex)
            {
                Log.Debug($"Error evaluating variable {variableName}: {ex.Message}");
                xlCell.Value = $"Error: {ex.Message}";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // 값이 null이면 처리하지 않음
            if (value == null)
            {
                xlCell.SetValue(Blank.Value);
                return;
            }

            // 포맷 적용 시도
            try
            {
                if (value is DateTime dateTime)
                {
                    // 날짜 값 특별 처리: 값 설정 전에 스타일 설정
                    if (formatString == "d")
                    {
                        // 단순 날짜 형식(Short Date)
                        xlCell.Style.DateFormat.Format = formatString;
                        xlCell.Value = dateTime.Date; // 시간 부분 제거
                    }
                    else
                    {
                        // 다른 날짜 형식들
                        xlCell.Style.DateFormat.Format = formatString;
                        xlCell.Value = dateTime;
                    }

                    Log.Debug($"Applied date format: {formatString} to {dateTime}");
                }
                else if (value is IFormattable formattable)
                {
                    try
                    {
                        // 기본 .NET 포맷팅 시도
                        var formattedValue = formattable.ToString(formatString, CultureInfo.CurrentCulture);
                        xlCell.Value = formattedValue;
                    }
                    catch (FormatException)
                    {
                        // .NET 포맷팅 실패 시 Excel 포맷 적용
                        xlCell.SetValue(value);
                        xlCell.Style.NumberFormat.Format = formatString;
                    }
                }
                else
                {
                    // 포맷팅할 수 없는 값은 직접 설정
                    xlCell.SetValue(value);
                }

                Log.Debug($"Applied format: {formatString}");
            }
            catch (Exception ex)
            {
                Log.Debug($"Error applying format {formatString}: {ex.Message}");
                // 포맷 적용 실패 시 원본 값 설정
                xlCell.SetValue(value);
                // Comment 속성 대신 GetComment() 메서드 사용
                var comment = xlCell.GetComment();
                comment.AddText($"Format error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"General error in FormatTag: {ex.Message}");
            xlCell.Value = $"Error: {ex.Message}";
            xlCell.Style.Font.FontColor = XLColor.Red;
        }
    }
}