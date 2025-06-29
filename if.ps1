if (-not (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

$printers = Get-Printer -ComputerName srvprintproiz | ForEach-Object {
    $port = Get-PrinterPort -Name $_.PortName -ErrorAction SilentlyContinue
    
    [PSCustomObject]@{
        "InventoryNumber" = ""
        "Description" = ""
        "Owner" = ""
        "Оргтехника.Цвет_печати" = ""
        "Оргтехника.Производитель" = ($_.DriverName -split ' ')[0] 
        "Оргтехника.Наименование" = $_.Name
        "Оргтехника.Формат_Печати" = ""
        "Оргтехника.Технология_Печати" = ""
        "Оргтехника.Серийный_номер" = ""
        "Оргтехника.Тип_подключения" = if ($port.PrinterHostAddress) { "Сетевой" } else { "Локальный" }
        "Оргтехника.Формат_сканирования" = ""
        "Оргтехника.Тип_картриджа" = ""
        "Оргтехника.IP-адрес" = $port.PrinterHostAddress
        "Оргтехника.Тип" = "Принтер"
        "Оргтехника.Модель" = $_.DriverName
        "Оргтехника.Здание" = ""
        "Оргтехника.Помещение" = ""
    }
}

$printers | Export-Excel -Path "C:\Printers_Inventory.xls" -WorksheetName "Принтеры" -AutoSize -FreezeTopRow -BoldTopRow