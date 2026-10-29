let
    Source = Excel.Workbook(File.Contents("C:\YOUR_FOLDER\bike-stations-sharing-data.xlsx"), null, true),
    Feuil1 = Source{[Item="Feuil1",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Feuil1, [PromoteAllScalars=true]),
    RenamedColumns = Table.RenameColumns(PromotedHeaders,{
        {"number","StationNumber"},{"name","StationName"},{"address","Address"},
        {"position","Position"},{"banking","Banking"},{"bonus","Bonus"},
        {"status","Status"},{"Contract Name","ContractName"},{"Bike Stands","BikeStands"},
        {"Available Bike Stands","AvailableBikeStands"},{"Available Bikes","AvailableBikes"},
        {"Last Update","LastUpdate"}
    }),
    TrimText = Table.TransformColumns(RenamedColumns,{
        {"StationName", each if _ is text then Text.Trim(_) else _, type text},
        {"Address", each if _ is text then Text.Trim(_) else _, type text},
        {"Status", each if _ is text then Text.Upper(Text.Trim(_)) else _, type text},
        {"ContractName", each if _ is text then Text.Lower(Text.Trim(_)) else _, type text}
    }),
    SetTypes = Table.TransformColumnTypes(TrimText,{
        {"StationNumber", Int64.Type},{"StationName", type text},{"Address", type text},
        {"Position", type text},{"Banking", type logical},{"Bonus", type logical},
        {"Status", type text},{"ContractName", type text},{"BikeStands", Int64.Type},
        {"AvailableBikeStands", Int64.Type},{"AvailableBikes", Int64.Type},{"LastUpdate", type datetimezone}
    }),
    SplitPosition = Table.SplitColumn(SetTypes, "Position", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Latitude", "Longitude"}),
    CoordinateTypes = Table.TransformColumnTypes(SplitPosition,{{"Latitude", type number},{"Longitude", type number}}),
    AddOccupiedStands = Table.AddColumn(CoordinateTypes, "OccupiedStands", each [BikeStands] - [AvailableBikeStands], Int64.Type),
    AddCapacityCheck = Table.AddColumn(AddOccupiedStands, "CapacityCheck", each if [BikeStands] = [AvailableBikeStands] + [AvailableBikes] then "Balanced" else "Check", type text),
    RemoveDuplicates = Table.Distinct(AddCapacityCheck),
    FilterValidStationNumbers = Table.SelectRows(RemoveDuplicates, each [StationNumber] <> null)
in
    FilterValidStationNumbers