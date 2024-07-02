..\connect.ps1

$ListNames = @(
#    "CoffeeMachines"
   "CoffeeMachineInventory"
#    "CoffeeMachineOrders"
#    "CoffeeMachineOrderInvoice"
#    "CoffeeMachineOrderDetails"
)

foreach ($ListName in $ListNames) {

    $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($null -eq $List) {
        Write-Host "List $ListName does not exist."
        continue
    }    
    Remove-PnPList -Identity $ListName -Force
}

Write-Host "All lists deleted successfully."

Disconnect-PnPOnline