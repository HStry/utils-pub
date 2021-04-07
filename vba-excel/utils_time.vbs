Attribute VB_Name = "utils_time"
Option Explicit

Public Function EpochToSerial(epoch_timestamp As Double) As Double
    EpochToSerial = 25569 + (epoch_timestamp / 86400)
End Function

Public Function EpochToDate(epoch_timestamp As Double) As Long
    EpochToDate = Floor(EpochToSerial(epoch_timestamp))
End Function

Public Function EpochToTime(epoch_timestamp As Double) As Double
    EpochToTime = (epoch_timestamp Mod 86400) / 86400
End Function

Public Function IsoWeeknum(serial_timestamp As Double) As Double
    IsoWeeknum = Application.WorksheetFunctions.IsoWeekNum(serial_timestamp)
End Function

