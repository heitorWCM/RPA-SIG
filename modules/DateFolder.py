from dataclasses import dataclass
import datetime
import os

@dataclass
class DateRangeConfig:
    path: str
    initial_date: str
    final_date: str
    
    def create_folder(self):
        """Creates the folder structure for the path"""
        os.makedirs(self.path, exist_ok=True)
        return self

def DeterminaDataECaminho(base_folder: str, subfolder: str, start_day: int = 1) -> DateRangeConfig:
    """
    Creates a monthly folder structure and returns date range configuration.
    
    Args:
        base_folder: Base directory path (e.g., "C:\\temp")
        subfolder: subfolder/project name for the subfolder
        start_day: Day of month to consider as "start" (default: 3)
    
    Returns:
        DateRangeConfig with path, initial_date, and final_date
    """
    today = datetime.date.today()
    first_this_month = today.replace(day=start_day)
    
    # Last month range
    first_last_month = (first_this_month - datetime.timedelta(days=1)).replace(day=1)
    last_day_prev_month = first_this_month - datetime.timedelta(days=1)
    
    # Format dates for output
    initial_date = first_last_month.strftime("%d%m%Y")
    final_date = last_day_prev_month.strftime("%d%m%Y")
    
    # Create path
    path = os.path.join(
        base_folder,
        subfolder,
        first_this_month.strftime("%Y"),
        f"{first_this_month.strftime('%m')}.{first_this_month.strftime('%b').upper()}"
    ).replace("/", "\\")
    
    return DateRangeConfig(
        path=path,
        initial_date=initial_date,
        final_date=final_date
    )