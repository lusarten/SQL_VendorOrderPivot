/*Instructions to transfer the results into Excel
UPC Code must be 12 or 14 digits, to transfer the results correctly please, 
set column 'H' as text to show the leading '0's, 
Select column 'H', Right-click/ Format Cells/Text/Ok
Copy the table and paste into the Excel File
*/


-- Check if the temporary table exists and drop it if it does
IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE id = OBJECT_ID('tempdb..#tempPivot')) 
BEGIN 
    DROP TABLE #tempPivot 
END

-- Create and populate the temporary table with the necessary data
SELECT DISTINCT
    w.VendProd AS [MM Stock #],  -- Vendor product number
    GETDATE() AS [Current Price As Of],  -- Current system date and time
    w.ListPrice AS [Current List Price],  -- Current List Price
    w.ReplCost AS [Current Net Price],  -- Current replacement cost
    t.RequestDate AS [New Price Effective Date],  -- New price effective date
    t2.ListPrice AS [New List Price],  -- New List Price
    t2.NetPrice AS [New Net Price],  -- New Net Price
    p1.UnitStock AS [Sales UOM],  -- Sales UOM
    
    -- UnitCode from MMPriceServiceUomDiscount
    t.UnitCode AS [UnitCode],

    -- MM Minimum Order UOM from MMPriceServiceUomDiscount
    t2.SalesUnitCode AS [MM Minimum Order UOM],  -- This is the new UOM field

    -- MM Minimum Order Qty from MMPriceService
    t.MinOrderQty AS [MM Minimum Order Qty],

    -- MM Incremental Order Qty from MMPriceService
    t.MinOrderInc AS [MM Incremental Order Qty],
    
    -- Conversion for sorting
    t.Conversion AS [Conversion],  -- Used for ordering, but won't show in the report

    -- Order Quantity for pivoting
    t.OrderQuantity AS [OrderQuantity],

    -- Discount Percentage for pivoting
    t.DiscountPercentage AS [DiscountPercentage],

    -- Net Price for pivoting
    t.NetPrice AS [NetPrice],
    
    -- Sales UPC - must be 14 digits
    CASE 
        WHEN LEN(
            CASE 
                WHEN sv.section2 = '00' THEN 
                    RIGHT('000000' + CONVERT(varchar,(CONVERT(int,sv.Section3))),6) + 
                    RIGHT('00000' + CONVERT(varchar,(CONVERT(int,sv.Section4))),5) + 
                    LEFT(CONVERT(varchar,(CONVERT(int,sv.Section5))),1) 
                ELSE 
                    RIGHT('00' + CONVERT(varchar,(CONVERT(int,sv.section2))),2) + 
                    RIGHT('000000' + CONVERT(varchar,(CONVERT(int,sv.Section3))),6) + 
                    RIGHT('00000' + CONVERT(varchar,(CONVERT(int,sv.Section4))),5) + 
                    LEFT(CONVERT(varchar,(CONVERT(int,sv.Section5))),1) 
            END) < 14 
        THEN 
            RIGHT('00000000000000' + 
            CASE 
                WHEN sv.section2 = '00' THEN 
                    RIGHT('000000' + CONVERT(varchar,(CONVERT(int,sv.Section3))),6) + 
                    RIGHT('00000' + CONVERT(varchar,(CONVERT(int,sv.Section4))),5) + 
                    LEFT(CONVERT(varchar,(CONVERT(int,sv.Section5))),1) 
                ELSE 
                    RIGHT('00' + CONVERT(varchar,(CONVERT(int,sv.section2))),2) + 
                    RIGHT('000000' + CONVERT(varchar,(CONVERT(int,sv.Section3))),6) + 
                    RIGHT('00000' + CONVERT(varchar,(CONVERT(int,sv.Section4))),5) + 
                    LEFT(CONVERT(varchar,(CONVERT(int,sv.Section5))),1) 
            END, 14)
        ELSE 
            RIGHT(
            CASE 
                WHEN sv.section2 = '00' THEN 
                    RIGHT('000000' + CONVERT(varchar,(CONVERT(int,sv.Section3))),6) + 
                    RIGHT('00000' + CONVERT(varchar,(CONVERT(int,sv.Section4))),5) + 
                    LEFT(CONVERT(varchar,(CONVERT(int,sv.Section5))),1) 
                ELSE 
                    RIGHT('00' + CONVERT(varchar,(CONVERT(int,sv.section2))),2) + 
                    RIGHT('000000' + CONVERT(varchar,(CONVERT(int,sv.Section3))),6) + 
                    RIGHT('00000' + CONVERT(varchar,(CONVERT(int,sv.Section4))),5) + 
                    LEFT(CONVERT(varchar,(CONVERT(int,sv.Section5))),1) 
            END, 14)
    END AS 'Sales UPC',
    
    -- UOM Descriptions
    CASE 
        WHEN p1.UnitStock IN ('AS', 'AST') THEN 'Assortment'
        WHEN p1.UnitStock IN ('BA', 'BAG') THEN 'Bale'
        WHEN p1.UnitStock IN ('BG', 'BAG') THEN 'Bag'
        WHEN p1.UnitStock IN ('BX', 'BOX') THEN 'Box'
        WHEN p1.UnitStock IN ('CY', 'CL') THEN 'Cylinder'
        WHEN p1.UnitStock IN ('CN', 'CAN') THEN 'Can'
        WHEN p1.UnitStock IN ('CS', 'CSD') THEN 'Case'
        WHEN p1.UnitStock IN ('CT', 'CTN') THEN 'Carton'
        WHEN p1.UnitStock IN ('CQ', 'CTG') THEN 'Cartridge'
        WHEN p1.UnitStock IN ('EA', 'BIT') THEN 'Each'
        WHEN p1.UnitStock IN ('GA','GLL') THEN 'US gallon'
        WHEN p1.UnitStock IN ('KT', 'KIT') THEN 'Kit'
        WHEN p1.UnitStock IN ('PK', 'PAK') THEN 'Pack'
        WHEN p1.UnitStock IN ('RO', 'RL') THEN 'Roll'
        WHEN p1.UnitStock = 'BDL' THEN 'Bundle'
        WHEN p1.UnitStock = 'BO' THEN 'Bottle'
        WHEN p1.UnitStock = 'BT' THEN 'Belt'
        WHEN p1.UnitStock = 'C62' THEN 'One'
        WHEN p1.UnitStock = 'CA' THEN 'Canister'
        WHEN p1.UnitStock = 'CD' THEN 'Card'
        WHEN p1.UnitStock = 'DSP' THEN 'Display'
        WHEN p1.UnitStock = 'DC' THEN 'Disk(Disc)'
        WHEN p1.UnitStock = 'DR' THEN 'Drum'
        WHEN p1.UnitStock = 'FOT' THEN 'Feet'
        WHEN p1.UnitStock = 'GRM' THEN 'Gram'
        WHEN p1.UnitStock = 'H87' THEN 'Piece'
        WHEN p1.UnitStock = 'KGM' THEN 'Kilogram'
        WHEN p1.UnitStock = 'LB' THEN 'Pound'
        WHEN p1.UnitStock = 'LBR' THEN 'US pound'
        WHEN p1.UnitStock = 'MTK' THEN 'Square meter'
        WHEN p1.UnitStock = 'MTR' THEN 'Meter'
        WHEN p1.UnitStock = 'PA' THEN 'Pail'
        WHEN p1.UnitStock = 'PD' THEN 'Pad'
        WHEN p1.UnitStock = 'PC' THEN 'Piece'
        WHEN p1.UnitStock = 'PF' THEN 'Pallet'
        WHEN p1.UnitStock = 'PKG' THEN 'Package'
        WHEN p1.UnitStock = 'PL' THEN 'Pallet'
        WHEN p1.UnitStock = 'PR' THEN 'Pair'
        WHEN p1.UnitStock = 'TB' THEN 'Tube'
        WHEN p1.UnitStock = 'QTL' THEN 'Liquid quart (US)'
        WHEN p1.UnitStock = 'REL' THEN 'Reel'
        WHEN p1.UnitStock = 'ROL' THEN 'Roll'     
        WHEN p1.UnitStock = 'SET' THEN 'Set'
        WHEN p1.UnitStock = 'ST' THEN 'Sheet'
        WHEN p1.UnitStock = 'WH' THEN 'Wheel'
        WHEN p1.UnitStock = 'YRD' THEN 'Yard'
        ELSE 'Unknown'
    END AS [UOM Description],
    ROW_NUMBER() OVER (PARTITION BY w.VendProd ORDER BY t.Conversion) AS QuantityLevel
INTO #tempPivot
FROM 
    DataWarehouse.Base.icsw w
    JOIN DataWarehouse.base.icsp p1 
        ON p1.Prod = w.Prod  
        AND p1.MetaDeletedInd = 'n'  
        AND p1.MetaEndEffDt = '9998/12/31'  
        AND p1.StatusType = 'a'  
        AND w.Whse = 'MUSK'  -- Warehouse MUSK is a must
        AND w.PriceType NOT IN ('MTGS', 'KITR')  -- Exclude Kits
    LEFT JOIN [Transportation].[mm].[MMPriceServiceUomDiscount] t
        ON t.mmStockNumber = w.VendProd  
    LEFT JOIN [Transportation].[mm].[MMPriceService] t2
        ON t2.mmStockNumber = w.VendProd  
    LEFT JOIN [DataWarehouse].[base].[icsv] sv  
        ON sv.Prod = w.Prod  
    LEFT JOIN [DataWarehouse].[base].[zicspaux] x  
        ON x.prod = w.prod
        AND x.metaendeffdt = '12/31/9998'
        AND x.MetaDeletedInd = 'n'
WHERE 
    w.MetaEndEffDt = '9998/12/31'  -- Only include active records
    AND w.MetaDeletedInd = 'n'  -- Exclude deleted records
    AND w.ArpVendNo IN ('123456', '765432')  -- Restrict to MM
    AND w.StatusType <> 'x';  -- Exclude inactive products

-- Pivot the data from the temporary table
SELECT 
    [MM Stock #],
    MAX([Current Price As Of]) AS [Current Price As Of],
    MAX([Current List Price]) AS [Current List Price],
    MAX([Current Net Price]) AS [Current Net Price],
    MAX([New Price Effective Date]) AS [New Price Effective Date],
    MAX([New List Price]) AS [New List Price],
    MAX([New Net Price]) AS [New Net Price],
    MAX([Sales UPC]) AS [Sales UPC],
    MAX([Sales UOM]) AS [Sales UOM],
    MAX([UOM Description]) AS [UOM Description],
    MAX([MM Minimum Order UOM]) AS [MM Minimum Order UOM], 
    MAX([MM Minimum Order Qty]) AS [MM Minimum Order Qty],
    MAX([MM Incremental Order Qty]) AS [MM Incremental Order Qty],
    ISNULL(MAX(CONVERT(VARCHAR, [1])), '') AS [Quantity Discount (Level 1) - Qty],
    ISNULL(MAX([UnitCode1]), '') AS [Quantity Discount (Level 1) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage1])), '') AS [Quantity % Discount (Level 1)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice1])), '') AS [Quantity Discount (Level 1) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [2])), '') AS [Quantity Discount (Level 2) - Qty],
    ISNULL(MAX([UnitCode2]), '') AS [Quantity Discount (Level 2) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage2])), '') AS [Quantity % Discount (Level 2)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice2])), '') AS [Quantity Discount (Level 2) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [3])), '') AS [Quantity Discount (Level 3) - Qty],
    ISNULL(MAX([UnitCode3]), '') AS [Quantity Discount (Level 3) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage3])), '') AS [Quantity % Discount (Level 3)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice3])), '') AS [Quantity Discount (Level 3) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [4])), '') AS [Quantity Discount (Level 4) - Qty],
    ISNULL(MAX([UnitCode4]), '') AS [Quantity Discount (Level 4) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage4])), '') AS [Quantity % Discount (Level 4)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice4])), '') AS [Quantity Discount (Level 4) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [5])), '') AS [Quantity Discount (Level 5) - Qty],
    ISNULL(MAX([UnitCode5]), '') AS [Quantity Discount (Level 5) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage5])), '') AS [Quantity % Discount (Level 5)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice5])), '') AS [Quantity Discount (Level 5) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [6])), '') AS [Quantity Discount (Level 6) - Qty],
    ISNULL(MAX([UnitCode6]), '') AS [Quantity Discount (Level 6) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage6])), '') AS [Quantity % Discount (Level 6)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice6])), '') AS [Quantity Discount (Level 6) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [7])), '') AS [Quantity Discount (Level 7) - Qty],
    ISNULL(MAX([UnitCode7]), '') AS [Quantity Discount (Level 7) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage7])), '') AS [Quantity % Discount (Level 7)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice7])), '') AS [Quantity Discount (Level 7) - New Net Unit Price],
    ISNULL(MAX(CONVERT(VARCHAR, [8])), '') AS [Quantity Discount (Level 8) - Qty],
    ISNULL(MAX([UnitCode8]), '') AS [Quantity Discount (Level 8) - UOM],
    ISNULL(MAX(CONVERT(VARCHAR, [DiscountPercentage8])), '') AS [Quantity % Discount (Level 8)],
    ISNULL(MAX(CONVERT(VARCHAR, [NetPrice8])), '') AS [Quantity Discount (Level 8) - New Net Unit Price]
FROM 
(
    SELECT 
        [MM Stock #],
        [Current Price As Of],
        [Current List Price],
        [Current Net Price],
        [New Price Effective Date],
        [New List Price],
        [New Net Price],
        [Sales UPC],
        [Sales UOM],
        [UOM Description],
        [UnitCode],
        [DiscountPercentage],
        [NetPrice],
        [MM Minimum Order UOM], 
        [MM Minimum Order Qty],
        [MM Incremental Order Qty],
        [OrderQuantity],
        QuantityLevel,
        [Conversion],
        CASE WHEN QuantityLevel = 1 THEN UnitCode END AS UnitCode1,
        CASE WHEN QuantityLevel = 1 THEN DiscountPercentage END AS DiscountPercentage1,
        CASE WHEN QuantityLevel = 1 THEN NetPrice END AS NetPrice1,
        CASE WHEN QuantityLevel = 2 THEN UnitCode END AS UnitCode2,
        CASE WHEN QuantityLevel = 2 THEN DiscountPercentage END AS DiscountPercentage2,
        CASE WHEN QuantityLevel = 2 THEN NetPrice END AS NetPrice2,
        CASE WHEN QuantityLevel = 3 THEN UnitCode END AS UnitCode3,
        CASE WHEN QuantityLevel = 3 THEN DiscountPercentage END AS DiscountPercentage3,
        CASE WHEN QuantityLevel = 3 THEN NetPrice END AS NetPrice3,
        CASE WHEN QuantityLevel = 4 THEN UnitCode END AS UnitCode4,
        CASE WHEN QuantityLevel = 4 THEN DiscountPercentage END AS DiscountPercentage4,
        CASE WHEN QuantityLevel = 4 THEN NetPrice END AS NetPrice4,
        CASE WHEN QuantityLevel = 5 THEN UnitCode END AS UnitCode5,
        CASE WHEN QuantityLevel = 5 THEN DiscountPercentage END AS DiscountPercentage5,
        CASE WHEN QuantityLevel = 5 THEN NetPrice END AS NetPrice5,
        CASE WHEN QuantityLevel = 6 THEN UnitCode END AS UnitCode6,
        CASE WHEN QuantityLevel = 6 THEN DiscountPercentage END AS DiscountPercentage6,
        CASE WHEN QuantityLevel = 6 THEN NetPrice END AS NetPrice6,
        CASE WHEN QuantityLevel = 7 THEN UnitCode END AS UnitCode7,
        CASE WHEN QuantityLevel = 7 THEN DiscountPercentage END AS DiscountPercentage7,
        CASE WHEN QuantityLevel = 7 THEN NetPrice END AS NetPrice7,
        CASE WHEN QuantityLevel = 8 THEN UnitCode END AS UnitCode8,
        CASE WHEN QuantityLevel = 8 THEN DiscountPercentage END AS DiscountPercentage8,
        CASE WHEN QuantityLevel = 8 THEN NetPrice END AS NetPrice8
    FROM #tempPivot
) AS SourceTable
PIVOT (
    MAX([OrderQuantity]) FOR QuantityLevel IN 
    (
        [1], [2], [3], [4], [5], [6], [7], [8]
    )
) AS PivotTable
GROUP BY 
    [MM Stock #]
HAVING 
    MAX([1]) IS NOT NULL  -- Eliminate rows where "Quantity Discount (Level 1) - Qty" is NULL
ORDER BY 
    [MM Stock #], 
    MAX([Conversion]); -- Order by the maximum Conversion within each group
