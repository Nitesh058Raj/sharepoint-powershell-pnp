# Import Required Modules
Import-Module ..\SharePoint-Lists.psm1

# Connect to a SharePoint site
..\connect.ps1


'''
-------------------------------------------------------------------------------
'''

# List Name: CoffeeMachines
$ListName = "CoffeeMachines"

# Fields
$Columns = @{
    Image_URL          = "Text"
    Price              = "Currency"
    Product_Type       = "Choice"
    Primary_Color      = "Text"
    Product_Summary    = "Note"
    Special_Feature    = "Text"
    Week_Avg_Cups      = "Text"
    Week_Avg_Espressos = "Text"
    Status             = "Choice"
}

# Choice Options
$Choice_Options = @{
    "Product_Type" = @("At Home Espresso Machine", "Commercial Espresso Machines", "Commercial Coffee Makers", "At Home Coffee Makers")
    "Status"       = @("Available", "Out of Stock")
}

Create-SharePoint-Fields -ListName $ListName -Columns $Columns -Choice_Options $Choice_Options

# Set the Title column to Name
Set-PnPField -List $ListName -Identity "Title" -Values @{Title = "Name" } 

'''
-------------------------------------------------------------------------------
'''


# List Name: CoffeeMachineOrders
$ListName = "CoffeeMachineOrders"

# Fields
$Columns = @{
    Total_Price = "Currency"
}

Create-SharePoint-Fields -ListName $ListName -Columns $Columns

'''
-------------------------------------------------------------------------------
'''


# Create a new List if not exists
$ListName = "CoffeeMachineOrderInvoice"

# Fields
$Columns = @{
    Order = "Lookup"
    File  = "Text"
}
              
# LookUp Attributes
$LookUpAttributes = @{
    "Order" = @("CoffeeMachineOrders", "Title")
}

Create-SharePoint-Fields -ListName $ListName -Columns $Columns -LookUpAttributes $LookUpAttributes

'''
-------------------------------------------------------------------------------
'''


# List Name: CoffeeMachineOrderDetails
$ListName = "CoffeeMachineOrderDetails"

# Fields
$Columns = @{
    Order      = "Lookup"
    Product    = "Lookup"
    Quantity   = "Number"
    Unit_Price = "Currency"
}
              
# LookUp Attributes 
$LookUpAttributes = @{
    "Order"   = @("CoffeeMachineOrders", "Title")
    "Product" = @("CoffeeMachines", "Name")
}

Create-SharePoint-Fields -ListName $ListName -Columns $Columns -LookUpAttributes $LookUpAttributes


'''
-------------------------------------------------------------------------------
'''

#List Name
$ListName = "CoffeeMachineInventory"

#Fields
$Columns = @{
    Product  = "Lookup"
    Quantity = "Number"
    Status   = "Choice"
}

#Choice Options
$Choice_Options = @{
    "Status" = @("Available", "Out of Stock")
}

#LookUp Attributes
$LookUpAttributes = @{
    "Product" = @("CoffeeMachines", "Title")
    "Status"  = @("CoffeeMachines", "Status")
}

Create-SharePoint-Fields -ListName $ListName -Columns $Columns -Choice_Options $Choice_Options -LookUpAttributes $LookUpAttributes


# Disconnect from SharePoint
Disconnect-PnPOnline