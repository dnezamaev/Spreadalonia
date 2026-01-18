using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace Demo;

public partial class DatePickerWindowViewModel : ObservableValidator
{
    public DatePickerWindowOptions Options { get; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsValid))]
    [NotifyDataErrorInfo]
    [Required(ErrorMessage = "Select date")]
    [CustomValidation(typeof(DatePickerWindowViewModel), nameof(ValidateDate))]
    public DateTimeOffset? selectedDate;

    public bool IsValid => !HasErrors;

    /// <summary>
    /// Notifies parent window that date is selected.
    /// </summary>
    public event EventHandler<DateTime>? DatePicked;

    public ICommand OkButtonPressedCommand { get; }


    public DatePickerWindowViewModel(DatePickerWindowOptions options)
    {
        Options = options;

        SelectedDate = Options.StartDate;

        OkButtonPressedCommand = new AsyncRelayCommand(OkButtonPressed);
    }

    private Task OkButtonPressed()
    {
        if (SelectedDate is null || HasErrors) return Task.CompletedTask;

        DatePicked?.Invoke(this, SelectedDate.Value.Date);

        return Task.CompletedTask;
    }

    public static ValidationResult ValidateDate(DateTimeOffset value, ValidationContext context)
    {
        var instance = (DatePickerWindowViewModel)context.ObjectInstance;

        var date = value.Date;
        var dateFormat = instance?.Options?.DateFormat;
        var regexPattern = instance?.Options?.RegexValidationPattern;

        // No pattern or date format. Any value is possible.
        if (regexPattern is null || dateFormat is null)
        {
            return ValidationResult.Success;
        }

        var isValid = ValidateDateByRegex(regexPattern, date, dateFormat);

        return isValid ? 
            ValidationResult.Success : 
            new ValidationResult("Incorrect date selected.");
    }

    public static bool ValidateDateByRegex(string regexPattern, DateTime date, string dateFormat)
    {
        var regex = new Regex(regexPattern);
        return regex.IsMatch(date.ToString(dateFormat));
    }
}

public class DatePickerWindowOptions
{
    public DateTime? StartDate { get; set; }

    public string? RegexValidationPattern { get; set; }

    public string? DateFormat { get; set; }
}