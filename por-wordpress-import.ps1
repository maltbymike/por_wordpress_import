#include settings file
. "$PSScriptRoot\settings.ps1"

$Query = "
    SELECT
        ItemFile.NUM AS SKU,
        NAME AS Name,
        TYPE AS por_type,
        QTY,
        DMG,
        Category, UserDefined1,
        INST,
        PER1, PER2, PER3, PER4, PER5, PER6, PER7, PER8, PER9, PER10,
        RATE1, RATE2, RATE3, RATE4, RATE5, RATE6, RATE7, RATE8, RATE9, RATE10,
        MANF,
        MODN,
        LOOKUP,
        Inactive,
        NoPrintOnContract,
        HideOnWebsite,
        Weight,
        Height,
        Width,
        Length
    FROM ItemFile
    WHERE TYPE = 'V' OR Header = '' AND TYPE IN ('T', 'A', 'D', 'H', 'K', 'U')"

$connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
$connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= $pathToDB"
$command = $connection.CreateCommand()
$command.CommandText = $Query
$adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet


$Query1 = " SELECT
                ItemKits.Num AS Num,
                ItemKits.ItemKey AS ItemKey,
                ItemFile.NUM AS KitItemNum
            FROM ItemKits
                LEFT JOIN ItemFile
                ON ItemKits.ItemKey = ItemFile.[KEY]
            WHERE
                ItemKits.DiscountPercent <> 100
            ORDER BY
                ItemKits.Num"
$command1 = $connection.CreateCommand()
$command1.CommandText = $Query1
$adapter1 = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command1
$dataset1 = New-Object -TypeName System.Data.DataSet
$adapter1.Fill( $dataset1 )

#Close connection to database
$connection.Close()

$totalProducts = $adapter.Fill($dataset)
$count = 1
$products = ForEach ($product in $dataset.Tables[0]) {
    #Add Columns for rental rates
    Add-Member -InputObject $product -NotePropertyName "Meta: _2_hour_rate" -NotePropertyValue $null
    Add-Member -InputObject $product -NotePropertyName "Meta: _4_hour_rate" -NotePropertyValue $null
    Add-Member -InputObject $product -NotePropertyName "Meta: _daily_rate" -NotePropertyValue $null
    Add-Member -InputObject $product -NotePropertyName "Meta: _weekly_rate" -NotePropertyValue $null
    Add-Member -InputObject $product -NotePropertyName "Meta: _4_week_rate" -NotePropertyValue $null

    #Add Columns for Woocommerce
    Add-Member -InputObject $product -NotePropertyName "Type" -NotePropertyValue "simple_rental"
    Add-Member -InputObject $product -NotePropertyName "Visibility in catalogue" -NotePropertyValue "hidden"
    Add-Member -InputObject $product -NotePropertyName "Published" -NotePropertyValue 0
    Add-Member -InputObject $product -NotePropertyName "Tax status" -NotePropertyValue "taxable"
    Add-Member -InputObject $product -NotePropertyName "Tax class" -NotePropertyValue "Rental Equipment"
    Add-Member -InputObject $product -NotePropertyName "Categories" -NotePropertyValue $null

    #Set 2 Hour Rate
    switch(2) {
        $product.PER1 { $product."Meta: _2_hour_rate" = $product.RATE1; break }
        $product.PER2 { $product."Meta: _2_hour_rate" = $product.RATE2; break }
        $product.PER3 { $product."Meta: _2_hour_rate" = $product.RATE3; break }
        $product.PER4 { $product."Meta: _2_hour_rate" = $product.RATE4; break }
        $product.PER5 { $product."Meta: _2_hour_rate" = $product.RATE5; break }
        $product.PER6 { $product."Meta: _2_hour_rate" = $product.RATE6; break }
        $product.PER7 { $product."Meta: _2_hour_rate" = $product.RATE7; break }
        $product.PER8 { $product."Meta: _2_hour_rate" = $product.RATE8; break }
    }

    #Set 4 Hour Rate
    switch(4) {
        $product.PER1 { $product."Meta: _4_hour_rate" = $product.RATE1; break }
        $product.PER2 { $product."Meta: _4_hour_rate" = $product.RATE2; break }
        $product.PER3 { $product."Meta: _4_hour_rate" = $product.RATE3; break }
        $product.PER4 { $product."Meta: _4_hour_rate" = $product.RATE4; break }
        $product.PER5 { $product."Meta: _4_hour_rate" = $product.RATE5; break }
        $product.PER6 { $product."Meta: _4_hour_rate" = $product.RATE6; break }
        $product.PER7 { $product."Meta: _4_hour_rate" = $product.RATE7; break }
        $product.PER8 { $product."Meta: _4_hour_rate" = $product.RATE8; break }
    }

    #Set Daily Rate
    switch(24) {
        $product.PER1 { $product."Meta: _daily_rate" = $product.RATE1; break }
        $product.PER2 { $product."Meta: _daily_rate" = $product.RATE2; break }
        $product.PER3 { $product."Meta: _daily_rate" = $product.RATE3; break }
        $product.PER4 { $product."Meta: _daily_rate" = $product.RATE4; break }
        $product.PER5 { $product."Meta: _daily_rate" = $product.RATE5; break }
        $product.PER6 { $product."Meta: _daily_rate" = $product.RATE6; break }
        $product.PER7 { $product."Meta: _daily_rate" = $product.RATE7; break }
        $product.PER8 { $product."Meta: _daily_rate" = $product.RATE8; break }
    }

    #Set Weekly Rate
    switch(168) {
        $product.PER1 { $product."Meta: _weekly_rate" = $product.RATE1; break }
        $product.PER2 { $product."Meta: _weekly_rate" = $product.RATE2; break }
        $product.PER3 { $product."Meta: _weekly_rate" = $product.RATE3; break }
        $product.PER4 { $product."Meta: _weekly_rate" = $product.RATE4; break }
        $product.PER5 { $product."Meta: _weekly_rate" = $product.RATE5; break }
        $product.PER6 { $product."Meta: _weekly_rate" = $product.RATE6; break }
        $product.PER7 { $product."Meta: _weekly_rate" = $product.RATE7; break }
        $product.PER8 { $product."Meta: _weekly_rate" = $product.RATE8; break }
    }

    #Set 4 Week Rate
    switch(672) {
        $product.PER1 { $product."Meta: _4_week_rate" = $product.RATE1; break }
        $product.PER2 { $product."Meta: _4_week_rate" = $product.RATE2; break }
        $product.PER3 { $product."Meta: _4_week_rate" = $product.RATE3; break }
        $product.PER4 { $product."Meta: _4_week_rate" = $product.RATE4; break }
        $product.PER5 { $product."Meta: _4_week_rate" = $product.RATE5; break }
        $product.PER6 { $product."Meta: _4_week_rate" = $product.RATE6; break }
        $product.PER7 { $product."Meta: _4_week_rate" = $product.RATE7; break }
        $product.PER8 { $product."Meta: _4_week_rate" = $product.RATE8; break }
    }

    #Set Category in Woocommerce format
    $_aerial = "Rental > Aerial and Access Equipment > "
    $_airtools = "Rental > Air Compressors and Tools > "
    $_compaction = "Rental > Compaction Equipment > "
    $_concrete = "Rental > Concrete and Masonry Equipment > "
    $_earthmoving = "Rental > Earthmoving & Excavation Equipment > "
    $_fans = "Rental > Fans\, Dehumidifiers & Air Quality Tools > "
    $_floor = "Rental > Floor & Tile Equipment > "
    $_heat = "Rental > Heaters > "
    $_fenceanddeck = "Rental > Fence and Deck Tools > "
    $_homereno = "Rental > Home Renovation\, Cleaning & Maintenance Equipment > "
    $_lawnandgarden = "Rental > Lawn\, Garden & Backyard Equipment > "
    $_materialhandling = "Rental > Material Handling Equipment > "
    $_metalworking = "Rental > Metal Working Tools > "
    $_saws = "Rental > Saws > "
    $_jobsite = "Rental > Jobsite Equipment > "
    $_wood = "Rental > Wood\, Tree & Stump Tools > "
    $_party = "Rental > Party\, Entertainment & Training Room > "

    function productCategoryNames($categoryNumber) {
    # Translate category number to Woocommerce taxonomy
        switch($categoryNumber) {
            48  { $returnCategories = "Accessories"; break }
            40  { $returnCategories = $_aerial + "Harnesses and Lanyards"; break }
            116 { $returnCategories = $_aerial + "Scissor Lifts\, Electric"; break }
            117 { $returnCategories = $_aerial + "Scissor Lifts\, Rough Terrain"; break }
            118 { $returnCategories = $_aerial + "Boom Lifts\, Towable"; break }
            119 { $returnCategories = $_aerial + "Boom Lifts\, 4x4"; break }
            26  { $returnCategories = $_aerial + "Ladders, " +
                                        $_homereno + "Ladders"; break }
            95  { $returnCategories = $_aerial + "Scaffolding\, Bakers, " +
                                        $_homereno + "Scaffolding\, Bakers"; break }
            36  { $returnCategories = $_aerial + "Scaffolding\, Standard, " +
                                        $_homereno + "Scaffolding\, Standard"; break }
            12  { $returnCategories = $_airtools + "Air Compressors, " +
                                        $_jobsite + "Air Compressors"; break }
            78  { $returnCategories = $_airtools + "Air Nailers and Staplers, " +
                                        $_jobsite + "Air Nailers and Staplers"; break }
            13  { $returnCategories = $_airtools + "Air Tools and Accessories, " +
                                        $_jobsite + "Air Tools and Accessories"; break }
            69  { $returnCategories = $_airtools + "Sandblast Pots, " +
                                        $_jobsite + "Sandblast Pots"; break }
            16  { $returnCategories = $_compaction + "Vibratory Plate Compactors"; break }
            120 { $returnCategories = $_compaction + "Vibratory Rammers (Jumping Jacks)"; break }
            121 { $returnCategories = $_compaction + "Rollers"; break }
            122 { $returnCategories = $_compaction + "Hand Tampers"; break }
            123 { $returnCategories = $_compaction + "Paver Rollers"; break }
            151 { $returnCategories = $_concrete + "Bolt Cutters, " +
                                        $_jobsite + "Bolt Cutters"; break }
            18  { $returnCategories = $_concrete + "Concrete and Brick Cutters, " +
                                        $_saws + "Concrete and Brick Saws"; break }
            79  { $returnCategories = $_concrete + "Concrete Breakers"; break }
            124 { $returnCategories = $_concrete + "Concrete Buggies and Wheel Barrows, " +
                                        $_materialhandling + "Concrete Buggies and Wheel Barrows, " +
                                        $_lawnandgarden + "Wheel Barrows and Buggies, "; break }
            104 { $returnCategories = $_concrete + "Concrete Drills, " +
                                        $_homereno + "Drills"; break }
            126 { $returnCategories = $_concrete + "Concrete Drills > Diamond Drill Bits, " +
                                        $_homereno + "Drills > Diamond Drill Bits"; break }
            127 { $returnCategories = $_concrete + "Concrete Drills > Spline Drive Bits, " +
                                        $_homereno + "Drills > Spline Drive Bits"; break }
            128 { $returnCategories = $_concrete + "Concrete Drills > SDS+ Bits, " +
                                        $_homereno + "Drills > SDS+ Bits"; break }
            129 { $returnCategories = $_concrete + 'Concrete Grinders'; break }
            113 { $returnCategories = $_concrete + 'Concrete Mixers'; break }
            96  { $returnCategories = $_concrete + "Concrete Trowels and Screeds"; break }
            130 { $returnCategories = $_concrete + 'Concrete Vibrators'; break }
            42  { $returnCategories = $_earthmoving + "Compact Excavators"; break }
            125 { $returnCategories = $_earthmoving + "Compact Excavators > Attachments for Excavators"; break }
            43  { $returnCategories = $_earthmoving + "Compact Track Loaders"; break }
            3   { $returnCategories = $_earthmoving + "Skid Steer Loaders"; break }
            5   { $returnCategories = $_earthmoving + "Skid Steer Loaders > Attachments for Skid Steer Loaders, " +
                                        $_earthmoving + "Compact Track Loaders > Attachments for Compact Track Loaders"; break }
            105 { $returnCategories = $_earthmoving + "Small Articulating Loaders"; break }
            6   { $returnCategories = $_earthmoving + "Mini Track Loaders"; break }
            7   { $returnCategories = $_earthmoving + "Mini Track Loaders > Attachments for Mini Track Loaders, " +
                                        $_earthmoving + "Small Articulating Loaders > Attachment for Small Articulating Loaders"; break }
            94  { $returnCategories = $_earthmoving + "Tractor Loader/Backhoes"; break }
            41  { $returnCategories = $_earthmoving + "Tractor Loader/Backhoes > 3pt Hitch Implements"; break }
            109 { $returnCategories = $_earthmoving + "Skid Steer Loaders > Attachments for Skid Steer Loaders > Auger Bits - Bobcat, " +
                                        $_earthmoving + "Mini Track Loaders > Attachments for Mini Track Loaders > Auger Bits - Bobcat, " +
                                        $_earthmoving + "Compact Excavators > Attachments for Excavators > Auger Bits - Bobcat, " +
                                        $_fenceanddeck + "Post Hole Augers and Diggers > Auger Bits - Bobcat"; break }
            87  { $returnCategories = $_earthmoving + "Trenchers, " +
                                        $_jobsite + "Trenchers"; break }
            132 { $returnCategories = $_fans + "Air Cleaners"; break }
            131 { $returnCategories = $_fans + "Dehumidifiers"; break }
            84  { $returnCategories = $_fans + "Fans"; break }
            83  { $returnCategories = $_floor + "Floor\, Carpet and Upholstery Cleaning, " +
                                        $_homereno + "Floor\, Carpet and Upholstery Cleaning"; break }
            81  { $returnCategories = $_floor + "Floor Sanders, " +
                                        $_fenceanddeck + "Sanders"; break }
            82  { $returnCategories = $_floor + "Floor Installation Tools"; break }
            80  { $returnCategories = $_floor + "Floor Removal Tools"; break }
            92  { $returnCategories = $_floor + "Tile Cutters and Saws"; break }
            23  { $returnCategories = $_heat + "Electric Heaters"; break }
            133 { $returnCategories = $_heat + "Diesel/Kerosene Direct Fired Heaters"; break }
            134 { $returnCategories = $_heat + "Diesel Indirect Fired Heaters"; break }
            150 { $returnCategories = $_heat + "Diesel Indirect Fired Heaters > Indirect Heater Accessories"; break }
            135 { $returnCategories = $_heat + "Diesel Infrared Radiant Heaters"; break }
            136 { $returnCategories = $_heat + "Propane Patio Heaters"; break }
            14  { $returnCategories = $_homereno + "Automotive Tools"; break }
            138 { $returnCategories = $_homereno + "Automotive Tools, " +
                                        $_metalworking + "Grinders"; break }
            97  { $returnCategories = $_homereno + "Drain Cleaning and Inspection"; break }
            19  { $returnCategories = $_homereno + "Drills, " +
                                        $_metalworking + "Drills"; break }
            20  { $returnCategories = $_homereno + "Drywall Tools"; break }
            139 { $returnCategories = $_homereno + "Electrical Tools"; break }
            8   { $returnCategories = $_homereno + "Insulation Removal"; break }
            29  { $returnCategories = $_homereno + "Painting and Decorating"; break }
            31  { $returnCategories = $_homereno + "Plumbing Tools"; break }
            32  { $returnCategories = $_homereno + "Pressure Washers, " +
                                        $_jobsite + "Pressure Washers, " +
                                        $_fenceanddeck + "Pressure Washers"; break }
            93  { $returnCategories = $_homereno + "Roof\, Window and Siding Tools"; break }
            35  { $returnCategories = $_homereno + "Carpentry Saws, " +
                                        $_jobsite + "Carpentry Saws, " +
                                        $_saws + "Carpentry Saws, " +
                                        $_fenceanddeck + "Carpentry Saws, " +
                                        $_wood + "Carpentry Saws"; break }
            75  { $returnCategories = $_homereno + "Vertical Shore Posts, " +
                                        $_jobsite + "Vertical Shore Posts"; break }
            22  { $returnCategories = $_jobsite + "Generators"; break }
            89  { $returnCategories = $_jobsite + "Levels and Survey Equipment"; break }
            114 { $returnCategories = $_jobsite + 'Metal Detectors, ' +
                                        $_fenceanddeck + "Metal Detectors"; break }
            88  { $returnCategories = $_jobsite + "Lighting Equipment"; break }
            33  { $returnCategories = $_jobsite + "Water Pumps"; break }
            142 { $returnCategories = $_jobsite + "Water Pumps > Hoses"; break }
            53  { $returnCategories = $_jobsite + "Temporary Fencing"; break }
            108 { $returnCategories = $_fenceanddeck + "Post Hole Augers and Diggers"; break }
            110 { $returnCategories = $_fenceanddeck + 'Post Hole Augers and Diggers > Auger Bits - 1-3/8" Hex'; break }
            85  { $returnCategories = $_fenceanddeck + "Post Drivers"; break }
            111 { $returnCategories = $_fenceanddeck + 'Post Pullers'; break }
            112 { $returnCategories = $_fenceanddeck + 'Fence and Deck Installation Tools'; break }
            99  { $returnCategories = $_lawnandgarden + "Grass and Week Trimmers"; break }
            100 { $returnCategories = $_lawnandgarden + "Hedge Trimmers"; break }
            27  { $returnCategories = $_lawnandgarden + "Lawn and Garden Tools"; break }
            101 { $returnCategories = $_lawnandgarden + "Lawn Rollers"; break }
            98  { $returnCategories = $_lawnandgarden + "Leaf Blowers"; break }
            102 { $returnCategories = $_lawnandgarden + "Rototillers"; break }
            90  { $returnCategories = $_lawnandgarden + "Stihl Kombisystem Tools"; break }
            91  { $returnCategories = $_lawnandgarden + "Stihl Yardboss Tools"; break }
            103 { $returnCategories = $_lawnandgarden + "Sweepers"; break }
            28  { $returnCategories = $_materialhandling + "Carts\, Dollies and Moving Items"; break }
            140 { $returnCategories = $_materialhandling + "Material Lifts and Hoists"; break }
            141 { $returnCategories = $_materialhandling + "Jacks"; break }
            115 { $returnCategories = $_metalworking + "Metal Cutting Saws, " +
                                       $_saws + "Metal Cutting Saws"; break }
            86  { $returnCategories = $_wood + "Branch Cutters and Log Tools"; break }
            15  { $returnCategories = $_wood + "Chainsaws, " +
                                        $_lawnandgarden + "Chainsaws, " +
                                        $_saws + "Chainsaws"; break }
            144 { $returnCategories = $_wood + "Stump Grinders"; break }
            145 { $returnCategories = $_wood + "Wood Chippers"; break }
            146 { $returnCategories = $_wood + "Wood Splitters"; break }
            30  { $returnCategories = $_party + "Food and Beverage"; break }
            148 { $returnCategories = $_party + "Games and Raffles"; break }
            149 { $returnCategories = $_party + "Tables and Chairs"; break }
        }
        return $returnCategories
    }

    # Get category name from category number
    $product.Categories = productCategoryNames($product.Category)

    # Get additional category from POR UserDefined1 Field
    $additionalCategories = $product.UserDefined1.Split(",")
    if ($additionalCategories) {
        ForEach ($additionalCategory in $additionalCategories) {
            $product.Categories += ", "
            $product.Categories += productCategoryNames($additionalCategory.Trim())
        }
    }

    #Put Rental-Accessories in a hidden category
    if ($product.por_type -eq "A") { $product.Categories = "Accessories" }

    #Convert name field to title case
    $TextInfo = (Get-Culture).TextInfo
    $product.Name = $TextInfo.ToTitleCase( $TextInfo.ToLower($product.Name) )

    #capitalize all word with mixed numbers and letters eg 55XA, GS461
    #capitalize all KM-* words and MM-* words
    $name_parts = -split $product.Name
    $name_parts_working = @()
    foreach ($name_part in $name_parts) {
        if ($name_part -match '\d') { $name_part = $name_part.ToUpper() }
        if ($name_part -match 'KM-') { $name_part = $name_part.ToUpper() }
        if ($name_part -match 'MM-') { $name_part = $name_part.ToUpper() }
        $name_parts_working += $name_part
    }
    $product.Name = $name_parts_working -join " "


    #Set list of specail capitalization rules
    $SpecialCapitalization = @(
        'SDS',
        'SPS',
        'cfm',
        'hp',
        'lb',
        'psi',
        'w/',
        'ft/lb',
        'ft.lb.',
        'cu.ft.',
        'c/w',
        'ROC',
        '3pt',
        'btu/h',
        'kW',
        'btu',
        'ft',
        'w/AF',
        'Hz',
        'OD',
        'RAD')

    #only modify words that are complete ----- would like this to also match if it is beside a digit but adding digit search ends up removing the digit
    foreach ($key in $SpecialCapitalization) {
        $product.Name = $product.Name -ireplace "(\b)$key(\b)", $key
    }

    #Set Visibility
    if ($product.HideOnWebsite -eq 0) { $product."Visibility in catalogue" = "visible" }
    #if ($product.por_type -eq "A") { $product."Visibility in catalogue" = "search" }

    #Set products that are not marked inactive as Published
    if ($product.Inactive -eq 0) { $product.Published = 1 }

    #Get the current products Kit Items
    $upsells = @()
    $upsells = $dataset1.Tables[0].Select("Num='"+$product.Item( "SKU" )+"'", 'ItemKey ASC')

    #Convert Kit Items to comma seperated list as required by woocommerce
    $upsellList = ''
    if ($upsells.Length -ne 0) {
        ForEach ($upsell in $upsells) {
            $upsellList = $upsellList + ',' + $upsell.KitItemNum.Trim()
        }
        $upsellList = $upsellList.Substring(1)
    }

    #Add Upsell Column to $product
    Add-Member -InputObject $product -NotePropertyName "Upsells" -NotePropertyValue $upsellList

    #Trim whitespace from SKU
    $product.SKU = $product.SKU.Trim()

    #add $product to $modified so that it can be used in the export
    $product

    Write-Host $count "of" $totalProducts ": " $product.Name

    $count++
}


$products |
    Select-Object -Property * -ExcludeProperty Category, UserDefined1, PER??, PER?, RATE??, RATE?, por_type, Inactive, NoPrintOnContract, HideOnWebsite, RowError, RowState, Table, ItemArray, HasErrors |
    Export-Csv -Path $pathToCSVOutput -NoTypeInformation

Write-Output $products | Out-GridView
