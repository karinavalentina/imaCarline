
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
add-type -AssemblyName microsoft.VisualBasic
#Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser
#Install-Module -Name SQLServer -Scope CurrentUser
# =================================================================================================================================================
# START ----------- VARIABLES
# =================================================================================================================================================
$PbiRestApi = "https://api.powerbi.com/v1.0/myorg/"
$PbiConnectionAS = "powerbi://api.powerbi.com/v1.0/myorg/"
$DatePrefix = Get-Date -Format "dd/MM/yyyy HH:mm" 
$credentials = Get-Credential
$FileName = "AppFinale - MFA.ps1"
$FilePath = Get-ChildItem -Recurse -Filter $FileName | Select-Object -ExpandProperty DirectoryName
$OutFileName = "\Configuration.txt"
$configFilePath = Join-Path $FilePath $OutFileName
# =================================================================================================================================================
# END ----------- VARIABLES
# =================================================================================================================================================



# =================================================================================================================================================
# START ----------- FORM CONECTION POWER BI SERVICE
# =================================================================================================================================================
# Verificar si el archivo de configuración existe y no está vacío
if ((Test-Path $configFilePath) -and ((Get-Content $configFilePath -Raw).Length -ne 0)) {
    $formConf = New-Object System.Windows.Forms.Form
    $formConf.Text = "Use saved configuration?"
    $formConf.Size = New-Object System.Drawing.Size(380, 150)
    $formConf.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

    $buttonConfigYes           = New-Object System.Windows.Forms.Button
    $buttonConfigYes.Text      = "Yes"
    $buttonConfigYes.Location  = New-Object System.Drawing.Point(80, 50)
    $buttonConfigYes.ForeColor = [System.Drawing.Color]::Black
    $buttonConfigYes.Add_Click({
        $global:SelectedOption = "Yes"
        $formConf.Close()
})
    $formConf.Controls.Add($buttonConfigYes)


    $buttonConfigNo            = New-Object System.Windows.Forms.Button
    $buttonConfigNo.Text       = "No"
    $buttonConfigNo.Location   = New-Object System.Drawing.Point(200, 50)
    $buttonConfigNo.ForeColor  = [System.Drawing.Color]::Black
    $buttonConfigNo.Add_Click({
        $global:SelectedOption = "No"
        $formConf.Close()

})
$formConf.Controls.Add($buttonConfigNo)
$formConf.ShowDialog()
}
else {
    $global:SelectedOption = "No"
}

if($global:SelectedOption -eq "Yes"){
    $config = Get-Content $configFilePath | ConvertFrom-Json
}

else {

$formPBIS = New-Object System.Windows.Forms.Form
$formPBIS.Text = "Configuration parameters"
$formPBIS.Size = New-Object System.Drawing.Size(800, 520)
$formPBIS.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

$label__ = New-Object System.Windows.Forms.Label
$label__.Text = "Connection parameters to Power BI service"
$label__.Location = New-Object System.Drawing.Point(20, 30)
$label__.Width = 500
$label__.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($label__)


$labelConnection = New-Object System.Windows.Forms.Label
$labelConnection.Text = "Connection to Power BI Service"
$labelConnection.Location = New-Object System.Drawing.Point(20, 60)
$labelConnection.ForeColor = [System.Drawing.Color]::Black
$labelConnection.Width = 180
$formPBIS.Controls.Add($labelConnection)

$dropdownCon = New-Object System.Windows.Forms.ComboBox
$dropdownCon.Items.Add("User Credentials")
$dropdownCon.Items.Add("App Registration")
$dropdownCon.Location = New-Object System.Drawing.Point(210, 60)
$dropdownCon.Width = 300
$formPBIS.Controls.Add($dropdownCon)

$labelTenant = New-Object System.Windows.Forms.Label
$labelTenant.Text = "Tenant ID"
$labelTenant.Location = New-Object System.Drawing.Point(20, 90)
$labelTenant.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($labelTenant)

$textboxTenant = New-Object System.Windows.Forms.TextBox
$textboxTenant.Location = New-Object System.Drawing.Point(210, 90)
$textboxTenant.Width = 300
$tenantID = $textboxTenant
$formPBIS.Controls.Add($textboxTenant)

$label_ = New-Object System.Windows.Forms.Label
$label_.Text = "Database where the dimension and fact tables will be written"
$label_.Location = New-Object System.Drawing.Point(20, 130)
$label_.Width = 500
$label_.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($label_)

$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "SQL Server"
$label1.Location = New-Object System.Drawing.Point(20, 160)
$label1.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($label1)

$textbox1 = New-Object System.Windows.Forms.TextBox
$textbox1.Location = New-Object System.Drawing.Point(210, 160)
$textbox1.Width = 300
$Server = $textbox1
$formPBIS.Controls.Add($textbox1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Text = "SQL DataBase"
$label2.Location = New-Object System.Drawing.Point(20, 190)
$label2.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($label2)

$textbox2 = New-Object System.Windows.Forms.TextBox
$textbox2.Location = New-Object System.Drawing.Point(210, 190)
$textbox2.Width = 300
$Database = $textbox2
$formPBIS.Controls.Add($textbox2)

$labeldrop = New-Object System.Windows.Forms.Label
$labeldrop.Text = "Authentication"
$labeldrop.Location = New-Object System.Drawing.Point(20, 220)
$labeldrop.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($labeldrop)

$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Items.Add("Azure Active Directory")
$dropdown.Items.Add("SQL Server Authentication")
$dromdown.Items.Add("Azure Active Directory (MFA)")
$dropdown.Location = New-Object System.Drawing.Point(210, 220)
$dropdown.Width = 300
$formPBIS.Controls.Add($dropdown)

$label3 = New-Object System.Windows.Forms.Label
$label3.Text = "User Name"
$label3.Location = New-Object System.Drawing.Point(20, 250)
$label3.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($label3)

$textbox3 = New-Object System.Windows.Forms.TextBox
$textbox3.Location = New-Object System.Drawing.Point(210, 250)
$textbox3.Width = 300
$UserName = $textbox3
$formPBIS.Controls.Add($textbox3)

$label4 = New-Object System.Windows.Forms.Label
$label4.Text = "Password"
$label4.Location = New-Object System.Drawing.Point(20,280)
$label4.ForeColor = [System.Drawing.Color]::Black
$formPBIS.Controls.Add($label4)

$textbox4 = New-Object System.Windows.Forms.TextBox
$textbox4.Location = New-Object System.Drawing.Point(210, 280)
$textbox4.Width = 300
$textbox4.PasswordChar = "*"
$textbox4.UseSystemPasswordChar = $true
$Password = $textbox4
$formPBIS.Controls.Add($textbox4)


$buttonPBIS2 = New-Object System.Windows.Forms.Button
$buttonPBIS2.Text = "Next"
$buttonPBIS2.Location = New-Object System.Drawing.Point(560, 60)
$buttonPBIS2.ForeColor = [System.Drawing.Color]::Black
$buttonPBIS2.Add_Click({
$formPBIS.Close()
})
$formPBIS.Controls.Add($buttonPBIS2)

$formPBIS.ShowDialog()

# =================================================================================================================================================
# END ----------- VARIABLES
# =================================================================================================================================================

$formTables = New-Object System.Windows.Forms.Form
$formTables.Text = "Configuration parameters"
$formTables.Size = New-Object System.Drawing.Size(800, 520)
$formTables.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

$label__ = New-Object System.Windows.Forms.Label
$label__.Text = "Tables creation Parameters"
$label__.Location = New-Object System.Drawing.Point(20, 30)
$label__.Width = 500
$label__.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($label__)

$SchemaConnection = New-Object System.Windows.Forms.Label
$SchemaConnection.Text = "Schema"
$SchemaConnection.Location = New-Object System.Drawing.Point(20, 60)
$SchemaConnection.ForeColor = [System.Drawing.Color]::Black
$SchemaConnection.Width = 180
$formTables.Controls.Add($SchemaConnection)

$textboxSchemaConnection = New-Object System.Windows.Forms.TextBox
$textboxSchemaConnection.Location = New-Object System.Drawing.Point(210, 60)
$textboxSchemaConnection.Width = 300
$Schema = $textboxSchemaConnection
$formTables.Controls.Add($textboxSchemaConnection)

$labelWS = New-Object System.Windows.Forms.Label
$labelWS.Text = "Dimension Workspaces"
$labelWS.Location = New-Object System.Drawing.Point(20, 90)
$labelWS.Width = 180
$labelWS.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelWS)

$textboxWS = New-Object System.Windows.Forms.TextBox
$textboxWS.Location = New-Object System.Drawing.Point(210, 90)
$textboxWS.Width = 300
$tableDimensionWorkspaces = $textboxWS
$formTables.Controls.Add($textboxWS)

$labelDIMReports = New-Object System.Windows.Forms.Label
$labelDIMReports.Text = "Dimension Reports"
$labelDIMReports.Width = 180
$labelDIMReports.Location = New-Object System.Drawing.Point(20, 120)
$labelDIMReports.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDIMReports)

$textboxDIMReports = New-Object System.Windows.Forms.TextBox
$textboxDIMReports.Location = New-Object System.Drawing.Point(210, 120)
$textboxDIMReports.Width = 300
$tableDimensionReports = $textboxDIMReports
$formTables.Controls.Add($textboxDIMReports)

$labelDataset = New-Object System.Windows.Forms.Label
$labelDataset.Text = "Dimension Datasets"
$labelDataset.Width =180
$labelDataset.Location = New-Object System.Drawing.Point(20, 150)
$labelDataset.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDataset)

$textboxDataset = New-Object System.Windows.Forms.TextBox
$textboxDataset.Location = New-Object System.Drawing.Point(210, 150)
$textboxDataset.Width = 300
$tableDimensionDatasets =$textboxDataset
$formTables.Controls.Add($textboxDataset)

$labelDF = New-Object System.Windows.Forms.Label
$labelDF.Text = "Dimension Dataflows"
$labelDF.Width = 180
$labelDF.Location = New-Object System.Drawing.Point(20, 180)
$labelDF.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDF)

$textboxDataflows = New-Object System.Windows.Forms.TextBox
$textboxDataflows.Location = New-Object System.Drawing.Point(210, 180)
$textboxDataflows.Width = 300
$tableDimensionDataflows = $textboxDataflows
$formTables.Controls.Add($textboxDataflows)

$labelDSRef = New-Object System.Windows.Forms.Label
$labelDSRef.Text = "Dimension Dataset Refresh"
$labelDSRef.Width =180
$labelDSRef.Location = New-Object System.Drawing.Point(20, 210)
$labelDSRef.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDSRef)

$textboxDSRef = New-Object System.Windows.Forms.TextBox
$textboxDSRef.Location = New-Object System.Drawing.Point(210, 210)
$textboxDSRef.Width = 300
$tableDimensionDatasetRefresh = $textboxDSRef
$formTables.Controls.Add($textboxDSRef)

$labelDSour = New-Object System.Windows.Forms.Label
$labelDSour.Text = "Dimension Dataset Data Source"
$labelDSour.Width=180
$labelDSour.Location = New-Object System.Drawing.Point(20,240)
$labelDSour.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDSour)

$textboxDSour = New-Object System.Windows.Forms.TextBox
$textboxDSour.Location = New-Object System.Drawing.Point(210, 240)
$textboxDSour.Width = 300
$tableDimensionDataSource = $textboxDSour
$formTables.Controls.Add($textboxDSour)

$labelDSRefFact = New-Object System.Windows.Forms.Label
$labelDSRefFact.Text = "Fact Dataset Refresh"
$labelDSRefFact.Width = 180
$labelDSRefFact.Location = New-Object System.Drawing.Point(20,270)
$labelDSRefFact.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDSRefFact)

$textboxDSRefFact = New-Object System.Windows.Forms.TextBox
$textboxDSRefFact.Location = New-Object System.Drawing.Point(210, 270)
$textboxDSRefFact.Width = 300
$tableFactRefreshDataset = $textboxDSRefFact
$formTables.Controls.Add($textboxDSRefFact)

$labelDFRefFact = New-Object System.Windows.Forms.Label
$labelDFRefFact.Text = "Fact Dataflow Refresh"
$labelDFRefFact.Width =180
$labelDFRefFact.Location = New-Object System.Drawing.Point(20,300)
$labelDFRefFact.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDFRefFact)

$textboxDFRefFact = New-Object System.Windows.Forms.TextBox
$textboxDFRefFact.Location = New-Object System.Drawing.Point(210, 300)
$textboxDFRefFact.Width = 300
$tableFactRefreshDataflow = $textboxDFRefFact
$formTables.Controls.Add($textboxDFRefFact)

$labelDSMCode = New-Object System.Windows.Forms.Label
$labelDSMCode.Text = "Fact Dataset M Code"
$labelDSMCode.Width =180
$labelDSMCode.Location = New-Object System.Drawing.Point(20,330)
$labelDSMCode.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDSMCode)

$textboxDSMCode = New-Object System.Windows.Forms.TextBox
$textboxDSMCode.Location = New-Object System.Drawing.Point(210, 330)
$textboxDSMCode.Width = 300
$tableFactCODEM = $textboxDSMCode
$formTables.Controls.Add($textboxDSMCode)

$labelDSMeasures = New-Object System.Windows.Forms.Label
$labelDSMeasures.Text = "Fact Dataset Measures"
$labelDSMeasures.Width =180
$labelDSMeasures.Location = New-Object System.Drawing.Point(20,360)
$labelDSMeasures.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelDSMeasures)

$textboxDSMeasures = New-Object System.Windows.Forms.TextBox
$textboxDSMeasures.Location = New-Object System.Drawing.Point(210, 360)
$textboxDSMeasures.Width = 300
$tableDatabaseMeasure = $textboxDSMeasures
$formTables.Controls.Add($textboxDSMeasures)

$labelView = New-Object System.Windows.Forms.Label
$labelView.Text = "Fact Report Views"
$labelView.Width =180
$labelView.Location = New-Object System.Drawing.Point(20,390)
$labelView.ForeColor = [System.Drawing.Color]::Black
$formTables.Controls.Add($labelView)

$textboxView = New-Object System.Windows.Forms.TextBox
$textboxView.Location = New-Object System.Drawing.Point(210, 390)
$textboxView.Width = 300
$tableReportViews = $textboxView
$formTables.Controls.Add($textboxView)

$buttonPBIS3 = New-Object System.Windows.Forms.Button
$buttonPBIS3.Text = "Next"
$buttonPBIS3.Location = New-Object System.Drawing.Point(560, 60)
$buttonPBIS3.ForeColor = [System.Drawing.Color]::Black
$buttonPBIS3.Add_Click({
$formTables.Close()
})
$formTables.Controls.Add($buttonPBIS3)
$formTables.ShowDialog()


# =================================================================================================================================================
# END ----------- VARIABLES
# =================================================================================================================================================



$formSDB = New-Object System.Windows.Forms.Form
$formSDB.Text = "Configuration parameters"
$formSDB.Size = New-Object System.Drawing.Size(800, 520)
$formSDB.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

$label_ = New-Object System.Windows.Forms.Label
$label_.Text = "First database for table and column matching"
$label_.Location = New-Object System.Drawing.Point(20, 30)
$label_.Width = 500
$label_.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($label_)

$labelServer_1 = New-Object System.Windows.Forms.Label
$labelServer_1.Text = "SQL Server"
$labelServer_1.Location = New-Object System.Drawing.Point(20, 60)
$labelServer_1.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelServer_1)

$textboxServer_1 = New-Object System.Windows.Forms.TextBox
$textboxServer_1.Location = New-Object System.Drawing.Point(210, 60)
$textboxServer_1.Width = 300
$Server_1 = $textboxServer_1
$formSDB.Controls.Add($textboxServer_1)

$labelDB_1 = New-Object System.Windows.Forms.Label
$labelDB_1.Text = "SQL DataBase"
$labelDB_1.Location = New-Object System.Drawing.Point(20, 90)
$labelDB_1.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelDB_1)

$textboxDB_1 = New-Object System.Windows.Forms.TextBox
$textboxDB_1.Location = New-Object System.Drawing.Point(210, 90)
$textboxDB_1.Width = 300
$Database_1 = $textboxDB_1
$formSDB.Controls.Add($textboxDB_1)

$labeldropDB_1 = New-Object System.Windows.Forms.Label
$labeldropDB_1.Text = "Authentication"
$labeldropDB_1.Location = New-Object System.Drawing.Point(20, 120)
$labeldropDB_1.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labeldropDB_1)

$dropdownDB_1 = New-Object System.Windows.Forms.ComboBox
$dropdownDB_1.Items.Add("Azure Active Directory")
$dropdownDB_1.Items.Add("SQL Server Authentication")
$dropdownDB_1.Location = New-Object System.Drawing.Point(210, 120)
$dropdownDB_1.Width = 300
$formSDB.Controls.Add($dropdownDB_1)

$labelUS_1 = New-Object System.Windows.Forms.Label
$labelUS_1.Text = "User Name"
$labelUS_1.Location = New-Object System.Drawing.Point(20, 150)
$labelUS_1.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelUS_1)

$textboxUS_1 = New-Object System.Windows.Forms.TextBox
$textboxUS_1.Location = New-Object System.Drawing.Point(210, 150)
$textboxUS_1.Width = 300
$UserName_1 = $textboxUS_1
$formSDB.Controls.Add($textboxUS_1)

$labelP_1 = New-Object System.Windows.Forms.Label
$labelP_1.Text = "Password"
$labelP_1.Location = New-Object System.Drawing.Point(20,180)
$labelP_1.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelP_1)

$textboxP_1 = New-Object System.Windows.Forms.TextBox
$textboxP_1.Location = New-Object System.Drawing.Point(210, 180)
$textboxP_1.Width = 300
$textboxP_1.PasswordChar = "*"
$textboxP_1.UseSystemPasswordChar = $true
$Password_1 = $textboxP_1
$formSDB.Controls.Add($textboxP_1)

$label_ = New-Object System.Windows.Forms.Label
$label_.Text = "Second database for table and column matching"
$label_.Location = New-Object System.Drawing.Point(20, 230)
$label_.Width = 500
$label_.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($label_)

$labelServer_2 = New-Object System.Windows.Forms.Label
$labelServer_2.Text = "SQL Server"
$labelServer_2.Location = New-Object System.Drawing.Point(20, 260)
$labelServer_2.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelServer_2)

$textboxServer_2 = New-Object System.Windows.Forms.TextBox
$textboxServer_2.Location = New-Object System.Drawing.Point(210, 260)
$textboxServer_2.Width = 300
$Server_2 = $textboxServer_2
$formSDB.Controls.Add($textboxServer_2)

$labelDB_2 = New-Object System.Windows.Forms.Label
$labelDB_2.Text = "SQL DataBase"
$labelDB_2.Location = New-Object System.Drawing.Point(20, 290)
$labelDB_2.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelDB_2)

$textboxDB_2 = New-Object System.Windows.Forms.TextBox
$textboxDB_2.Location = New-Object System.Drawing.Point(210, 290)
$textboxDB_2.Width = 300
$Database_2 = $textboxDB_2
$formSDB.Controls.Add($textboxDB_2)

$labeldropDB_2 = New-Object System.Windows.Forms.Label
$labeldropDB_2.Text = "Authentication"
$labeldropDB_2.Location = New-Object System.Drawing.Point(20, 320)
$labeldropDB_2.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labeldropDB_2)

$dropdownDB_2 = New-Object System.Windows.Forms.ComboBox
$dropdownDB_2.Items.Add("Azure Active Directory")
$dropdownDB_2.Items.Add("SQL Server Authentication")
$dropdownDB_2.Location = New-Object System.Drawing.Point(210, 320)
$dropdownDB_2.Width = 300
$formSDB.Controls.Add($dropdownDB_2)

$labelUS_2 = New-Object System.Windows.Forms.Label
$labelUS_2.Text = "User Name"
$labelUS_2.Location = New-Object System.Drawing.Point(20, 350)
$labelUS_2.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelUS_2)

$textboxUS_2 = New-Object System.Windows.Forms.TextBox
$textboxUS_2.Location = New-Object System.Drawing.Point(210, 350)
$textboxUS_2.Width = 300
$UserName_2 = $textboxUS_2
$formSDB.Controls.Add($textboxUS_2)

$labelP_2 = New-Object System.Windows.Forms.Label
$labelP_2.Text = "Password"
$labelP_2.Location = New-Object System.Drawing.Point(20,380)
$labelP_2.ForeColor = [System.Drawing.Color]::Black
$formSDB.Controls.Add($labelP_2)

$textboxP_2 = New-Object System.Windows.Forms.TextBox
$textboxP_2.Location = New-Object System.Drawing.Point(210, 380)
$textboxP_2.Width = 300
$textboxP_2.PasswordChar = "*"
$textboxP_2.UseSystemPasswordChar = $true
$Password_2 = $textboxP_2
$formSDB.Controls.Add($textboxP_2)



$buttonConfig = New-Object System.Windows.Forms.Button
$buttonConfig.Text = "Save"
$buttonConfig.Location = New-Object System.Drawing.Point(560, 30)
$buttonConfig.ForeColor = [System.Drawing.Color]::Black
$buttonConfig.Add_Click({
$configuration = @{
    Server                       = $Server.Text
    Database                     = $Database.Text
    Authentication               = $dropdown.Text
    UserName                     = $UserName.Text
    Password                     = $Password.Text
    connectionPBI                = $dropdownCon.Text
    TenantID                     = $tenantID.Text
    Schema                       = $Schema.Text
    tableDimensionWorkspaces     = $tableDimensionWorkspaces.Text
    tableDimensionReports        = $tableDimensionReports.Text
    tableDimensionDatasets       = $tableDimensionDatasets.Text
    tableDimensionDataflows      = $tableDimensionDataflows.Text
    tableDimensionDatasetRefresh = $tableDimensionDatasetRefresh.Text
    tableDimensionDataSource     = $tableDimensionDataSource.Text
    tableFactRefreshDataset      = $tableFactRefreshDataset.Text
    tableFactRefreshDataflow     = $tableFactRefreshDataflow.Text
    tableFactCODEM               = $tableFactCODEM.Text
    tableDatabaseMeasure         = $tableDatabaseMeasure.Text
    tableReportViews             = $tableReportViews.Text
    Server_1                     = $Server_1.Text
    Database_1                   = $Database_1.Text
    UserName_1                   = $UserName_1.Text
    Authentication_1             = $dropdownDB_1.Text
    Password_1                   = $Password_1.Text
    Server_2                     = $Server_2.Text
    Database_2                   = $Database_2.Text
    UserName_2                   = $UserName_2.Text
    Authentication_2             = $dropdownDB_2.Text
    Password_2                   = $Password_2.Text
}

$configuration | ConvertTo-Json | Set-Content $configFilePath -Force
$formSDB.Close()

})
$formSDB.Controls.Add($buttonConfig)
$formSDB.ShowDialog()
}

$config = Get-Content $configFilePath | ConvertFrom-Json


if ($config.connectionPBI -eq "User Credentials")
{$connectionPBI = Connect-PowerBIServiceAccount -Credential $credentials}
else {$connectionPBI = Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credentials -TenantId $config.TenantID }

$connectionPBI

if ($config.Authentication -eq "Azure Active Directory")
{$Authentication = ";Authentication=Active Directory Password"}
else {$Authentication = ""}

# =================================================================================================================================================
# END ----------- FORM CONECTION POWER BI SERVICE
# =================================================================================================================================================

# =================================================================================================================================================
# START ----------- SQL TABLES
# =================================================================================================================================================
$Schema = $config.Schema
$tableDimensionWorkspaces     = "[$Schema].[$($config.tableDimensionWorkspaces)]"
$tableDimensionDatasets       = "[$Schema].[$($config.tableDimensionDatasets)]"
$tableDimensionReports        = "[$Schema].[$($config.tableDimensionReports)]"
$tableDimensionDataflows      = "[$Schema].[$($config.tableDimensionDataflows)]"
$tableDimensionDatasetRefresh = "[$Schema].[$($config.tableDimensionDatasetRefresh)]"
$tableDimensionDataSource     = "[$Schema].[$($config.tableDimensionDataSource)]"
$tableFactRefreshDataset      = "[$Schema].[$($config.tableFactRefreshDataset)]"
$tableFactRefreshDataflow     = "[$Schema].[$($config.tableFactRefreshDataflow)]"
$tableFactCODEM               = "[$Schema].[$($config.tableFactCODEM)]"
$tableDatabaseMeasure         = "[$Schema].[$($config.tableDatabaseMeasure)]" 
$tableReportViews             = "[$Schema].[$($config.tableReportViews)]"
         
# =================================================================================================================================================
# END ----------- SQL TABLES
# =================================================================================================================================================




# =================================================================================================================================================
# START ----------- FORM CONECTION SQL
# =================================================================================================================================================
$Server = $config.Server
$Database = $config.Database
$UserName= $config.UserName
$Password = $config.Password
# =================================================================================================================================================
# END ----------- FORM CONECTION SQL
# =================================================================================================================================================



# =================================================================================================================================================
# START ----------- CONECTION TO SQL
# =================================================================================================================================================
try {
    $connectionString = "Server=$Server;Database=$Database;User Id=$UserName;Password=$Password$Authentication"
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString
    $connection.Open()
    Write-Host $($connection.State)
} 
catch {
    Write-Host $_.Exception.Message
    Write-Host $($_.Exception.StackTrace)
}

# =================================================================================================================================================
# END ----------- CONECTION TO SQL
# =================================================================================================================================================


# =================================================================================================================================================
# START ----------- CREATE TO SCHEMA 
# =================================================================================================================================================
$query = "Create schema pbi"
$command = $connection.CreateCommand()
$command.CommandText = $query
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
$query_schema = "Create schema $Schema"
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_schema -AccessToken $AccessToken
# =================================================================================================================================================
# END ----------- CREATE TO SCHEMA 
# =================================================================================================================================================


# =================================================================================================================================================
# START ----------- TABLES CREATION
# =================================================================================================================================================

$query_Create_Dim_Reports = "CREATE TABLE $tableDimensionReports
(
	[Id_Report]       [varchar](255)  NULL,
	[DatasetName]     [varchar](255)  NULL,
	[Id_Dataset]      [varchar](255)  NULL,
	[Owner]           [varchar](255)  NULL,
	[ReportName]      [varchar](255)  NULL,
	[ReportWebUrl]    [varchar](2000) NULL,
	[Id_Workspace]    [varchar](255)  NULL,
	[WorkspaceName]   [varchar](255)  NULL,
    [ExtractionDate]  [varchar](255)  NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_Dim_Reports
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_Dim_Reports -AccessToken $AccessToken

$query_Create_Dim_Dataflows = "CREATE TABLE $tableDimensionDataflows
(
	[Id_Dataflow]         [varchar](255)  NULL,
	[DataflowName]        [varchar](255)  NULL,
	[DataflowDescription] [varchar](2000) NULL,
	[ConfiguredBy]        [varchar](255)  NULL,
	[Users]               [varchar](2000) NULL,
	[Id_Workspace]        [varchar](255)  NULL,
	[WorkspaceName]       [varchar](255)  NULL,
	[ExtractionDate]      [varchar](255)  NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_Dim_Dataflows
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_Dim_Dataflows -AccessToken $AccessToken

$query_Create_Dim_workspaces = "CREATE TABLE $tableDimensionWorkspaces
(
	[Id_Workspace]                [varchar](255) NULL,
	[WorkspaceName]               [varchar](255) NULL,
	[IsReadOnly]                  [varchar](10)  NULL,
	[isOnDedicatedCapacity]       [varchar](255) NULL,
	[CapacityId]                  [varchar](255) NULL,
	[defaultDatasetStorageFormat] [varchar](255) NULL,
    [ExtractionDate]              [varchar](255) NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_Dim_workspaces
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_Dim_workspaces -AccessToken $AccessToken


$query_Create_Dim_datasets = "CREATE TABLE $tableDimensionDatasets
(
	[Id_Dataset] [varchar](255) NULL,
	[Dataset_Name] [varchar](255) NULL,
	[WebUrl] [varchar](2000) NULL,
	[AddRowsAPIEnabled] [varchar](10) NULL,
	[ConfiguredBy] [varchar](255) NULL,
	[isOnPremGatewayRequired] [varchar](10) NULL,
	[TargetStorageMode] [varchar](255) NULL,
	[CreatedDate] [datetime] NULL,
	[Id_Workspace] [varchar](255) NULL,
	[WorkspaceName] [varchar](255) NULL,
	[ExtractionDate] [varchar](255) NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_Dim_datasets
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_Dim_datasets -AccessToken $AccessToken


$query_Create_Dim_datasets_Refresh = "CREATE TABLE $tableDimensionDatasetRefresh
(
	[Id_Dataset] [varchar](255) NULL,
	[DirectQuery] [varchar](10) NULL,
	[LocalTimeZoneId] [varchar](255) NULL,
	[Enable] [varchar](255) NULL,
	[Times] [varchar](1000) NULL,
	[Days] [varchar](1000) NULL,
	[Frecuency] [varchar](255) NULL,
	[Id_Workspace] [varchar](255) NULL,
	[ExtractionDate] [varchar](255) NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_Dim_datasets_Refresh
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_Dim_datasets_Refresh -AccessToken $AccessToken

$query_Create_Dim_dataSource = "CREATE TABLE $tableDimensionDataSource
(
	[Id_Dataset] [varchar](255) NULL,
	[DatasetName] [varchar](255) NULL,
	[DatasourceType] [varchar](255) NULL,
	[Details] [varchar](1000) NULL,
	[Refresh] [varchar](10) NULL,
	[Id_Report] [varchar](255) NULL,
	[ReportName] [varchar](255) NULL,
	[Id_Workspace] [varchar](255) NULL,
	[WorkspaceName] [varchar](255) NULL,
    [ExtractionDate] [varchar](255) NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_Dim_dataSource
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_Dim_dataSource -AccessToken $AccessToken



$query_Create_RefreshDataset = "CREATE TABLE $tableFactRefreshDataset

(
	[Transaccion_Id] [varchar](2000) NULL,
	[Id_Dataset] [varchar](255) NULL,
	[RefreshType] [varchar](255) NULL,
	[StartTime] [varchar](255) NULL,
	[EndTime] [varchar](255) NULL,
	[Status] [varchar](20) NULL,
	[Id_Workspace] [varchar](255) NULL,
	[WorkspaceName] [varchar](255) NULL,
	[ExtractionDate] [varchar](255) NULL,
	[serviceExceptionJson] [varchar](8000) NULL,
	[refreshAttempts] [varchar](8000) NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_RefreshDataset
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query $query_Create_RefreshDataset -AccessToken $AccessToken


$query_Create_RefreshDataflow = "CREATE TABLE $tableFactRefreshDataflow
(
	[Transaccion_Id] [varchar](2000) NULL,
	[Id_Dataflow] [varchar](255) NULL,
	[RefreshType] [varchar](255) NULL,
	[StartTime] [varchar](255) NULL,
	[EndTime] [varchar](255) NULL,
	[Status] [varchar](20) NULL,
	[Id_Workspace] [varchar](255) NULL,
	[WorkspaceName] [varchar](255) NULL,
	[ExtractionDate] [varchar](255) NULL
)"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_RefreshDataflow
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query  $query_Create_RefreshDataflow -AccessToken $AccessToken


$query_Create_tableMCODE = "CREATE TABLE $tableFactCODEM
(
	[ID] [int] NULL,
	[Id_Dataset] [varchar](255) NULL,
    [DatasetName] [varchar](1000) NULL,
    [Id_Workspace] [varchar](255) NULL,
	[M_Code] [varchar](8000) NULL,
	[LoadToReport] [varchar](10) NULL,
	[Table_Source] [varchar](255) NULL,
	[SQL_Server] [varchar](255) NULL,
	[SQL_Database] [varchar](255) NULL,
	[TableName(PBI)] [varchar](100) NULL,
	[TableName(SQL)] [varchar](7000) NULL,
	[Columns(SQL)] [varchar](7000) NULL,
    [ExtractionDate] [varchar](255) NULL,
    [TableName] [varchar](255) NULL,
	[ModifiedTime] [varchar](255) NULL,
    [StructureModifiedTime] [varchar](255) NULL
)
"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_tableMCODE
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query  $query_Create_tableMCODE -AccessToken $AccessToken



$query_Create_tableDatabaseMeasure = "CREATE TABLE $tableDatabaseMeasure
(
    [Id_Dataset] [varchar](255) NULL,
    [TableID] [varchar](1000) NULL,
    [MeasureName] [varchar](1000) NULL,
    [MeasureType] INT,
    [Expression] [varchar](8000) NULL,        
    [ModifiedTime]  [varchar](255) NULL,
    [StructureModifiedTime]  [varchar](255) NULL,
    [ErrorMessage] [varchar](8000) NULL,
    [ExtractionDate] [varchar](255) NULL
    )
"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_tableDatabaseMeasure
$command.ExecuteNonQuery()

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query  $query_Create_tableDatabaseMeasure -AccessToken $AccessToken


$query_Create_tableReportViews = "CREATE TABLE $tableReportViews
(
	[Id_Report] [varchar](255) NULL,
	[DatasetName]  [varchar](255) NULL,
	[Date] [varchar](255) NULL,
	[DistributionMethod] [varchar](36) NULL,
	[User] [varchar](255) NULL,
	[Id_Workspace] [varchar](255) NULL,
	[WorkspaceName] [varchar](255) NULL,
	[Report_Rank] [int] NULL
    )
"
$command = $connection.CreateCommand()
$command.CommandText = $query_Create_tableReportViews
$command.ExecuteNonQuery()

Connect-AzAccount -TenantId "dca2e9ff-6315-4d80-895a-208a5a5962a2"
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $database  -Username $UserName -Password $Password -Query  $query_Create_tableReportViews -AccessToken $AccessToken



# =================================================================================================================================================
# END ----------- TABLES CREATION
# =================================================================================================================================================



# =================================================================================================================================================
# START  -------- Objects Creation
# =================================================================================================================================================
$DIM_Reports            = @()
$DatiReports            = @()
$DIM_Datasets           = @()
$DIM_Dataset_Refresh    = @()
$DIM_Dataflows          = @()
$FAC_Dataflow_Refresh   = @()
$FAC_Dataset_Refresh    = @()
$DIM_Dataset_Refresh_DQ = @()
$DIM_Workspaces         = @()
$extractCODEM           = @()
$extractCODEM2          = @()
$extractCODEM3          = @()
$DatasetMeasure         = @()
$ReportViews            = @()
$DatasetUsers           = @()
# =================================================================================================================================================
# END  -------- Objects Creation
# =================================================================================================================================================

$workspaces = Get-PowerBIWorkspace
# =================================================================================================================================================
# START  -------- DIM REPORTS
# =================================================================================================================================================
foreach ($workspace in $workspaces) {

    $reports = Get-PowerBIReport -WorkspaceId $workspace.Id
    
    foreach ($report in $reports) {
        $dataset = Get-PowerBIDataset -WorkspaceId $workspace.Id -Id $report.DatasetId      
        $reportInfo = @{
            Id_Report             = $report.Id
            ReportName            = $report.Name
            DatasetName           = $dataset.Name
            Id_Dataset            = $report.DatasetId
            Owner                 = $dataset.ConfiguredBy
            ReportWebUrl          = $report.WebUrl 
            Id_Workspace          = $workspace.Id
            WorkspaceName         = $workspace.Name
            ReportDatasetEmbedUrl = $report.EmbedUrl
            ExtractionDate        = $DatePrefix
        }
      $DIM_Reports += $reportInfo
    }
}


# =================================================================================================================================================
# START  -------- DIM DATASOURCE
# =================================================================================================================================================

foreach ($workspace in $workspaces) {
    $reports = Get-PowerBIReport -WorkspaceId $workspace.Id
    foreach ($report in $reports) {
        $connections = Get-PowerBIDatasource -DatasetId $report.DatasetId -WorkspaceId $workspace.id
        $dataSetId = $report.DatasetId
        $dataset = Get-PowerBIDataset -WorkspaceId $workspace.Id -Id $dataSetId
        $Refresh = $dataset.IsRefreshable 


        if ($null -ne $connections -and (-not ($connections -is [System.Collections.IList]))) {
            $DatiReport = @{
                WorkspaceName      = $($workspace.Name)
                Id_Workspace       = $($workspace.Id)
                ReportName         = $($report.Name)
                Id_Report          = $($report.Id)
                Details            = $connections.ConnectionDetails
                DatasourceType     = $connections.DatasourceType
                DatasourceId       = $connections.DatasourceId
                Refresh            = $Refresh
                Datasetname        = $dataset.Name
                Id_Dataset         = $dataset.id
                ExtractionDate     = $DatePrefix
                }
            $DatiReports += $DatiReport
        }   
    }
}
# =================================================================================================================================================
#END  -------- DIM DATASOURCE
# =================================================================================================================================================

 

# =================================================================================================================================================
# START  -------- Conections and Parametries
# =================================================================================================================================================
$GetWorkspaces = $PbiRestApi + "groups"
$AllWorkspaces = Invoke-PowerBIRestMethod -Method GET -Url $GetWorkspaces | ConvertFrom-Json
# =================================================================================================================================================
# END -------- Conections and Parametries
# =================================================================================================================================================


# =================================================================================================================================================
# START -------- Workspace GET
# =================================================================================================================================================
foreach ($WorkspaceId in $AllWorkspaces.value.id) {
# Export data parameters
$GetWorkspaceApiCall = $PbiRestApi + "groups/" + $WorkspaceId
$WorkspaceInfo = Invoke-PowerBIRestMethod -Method GET -Url $GetWorkspaceApiCall | ConvertFrom-Json
$WorkspaceInfo | Add-Member -MemberType NoteProperty 'ExtractionDate' -Value $DatePrefix -Force
$DIM_Workspaces += $WorkspaceInfo
# =================================================================================================================================================
# END   -------- Workspace GET
# =================================================================================================================================================


# =================================================================================================================================================
# START -------- Dataflows GET
# =================================================================================================================================================
$GetDataflowsApiCall = $PbiRestApi + "groups/" + $WorkspaceId + "/dataflows"
$AllDataflows = Invoke-PowerBIRestMethod -Method GET -Url $GetDataflowsApiCall | ConvertFrom-Json
$ListAllDataflows = $AllDataflows.value
# =================================================================================================================================================
# END -------- Dataflows GET
# =================================================================================================================================================




# =================================================================================================================================================
# START -------- Dataflows DIMENSION TABLE
# =================================================================================================================================================
$ListAllDataflows | ForEach-Object{
    $_ | Add-Member -MemberType NoteProperty 'Id_Workspace'   -Value $WorkspaceId        -Force
    $_ | Add-Member -MemberType NoteProperty 'WorkspaceName'  -Value $WorkspaceInfo.name -Force
    $_ | Add-Member -MemberType NoteProperty 'ExtractionDate' -Value $DatePrefix         -Force
}
$DIM_Dataflows += $ListAllDataflows 
# =================================================================================================================================================
# END -------- Dataflows DIMENSION TABLE
# =================================================================================================================================================



# =================================================================================================================================================
# START -------- Function Get Dataflow Refresh 
# =================================================================================================================================================
Function GetDataflowRefreshResults {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true)][string]$DataflowId
    )
    $GetDataflowRefreshHistory = $PbiRestApi + "groups/" + $WorkspaceId + "/dataflows/" + $DataflowId + "/transactions"
    $DataflowRefreshHistory = Invoke-PowerBIRestMethod -Method GET -Url $GetDataflowRefreshHistory | ConvertFrom-Json
    return $DataflowRefreshHistory.value
}
# =================================================================================================================================================
# START -------- Function aplication to create FAC Table Dataflow Refresh
# =================================================================================================================================================
foreach($dataflow in $ListAllDataflows) {
    $DataflowHistories = GetDataflowRefreshResults -DataflowId $dataflow.objectId
    foreach($DataflowHistory in $DataflowHistories) {
        Add-Member -InputObject $DataflowHistory -NotePropertyName 'Id_Dataflow'    -NotePropertyValue $dataflow.objectId
        Add-Member -InputObject $DataflowHistory -NotePropertyName 'Id_Workspace'   -NotePropertyValue $WorkspaceId -Force
        Add-Member -InputObject $DataflowHistory -NotePropertyName 'WorkspaceName'  -NotePropertyValue $WorkspaceInfo.name -Force
        Add-Member -InputObject $DataflowHistory -NotePropertyName 'ExtractionDate' -NotePropertyValue $DatePrefix -Force
        $FAC_Dataflow_Refresh += $DataflowHistory
    }  
}

# =================================================================================================================================================
# END  -------- Function Get Dataflow Refresh and Function aplication to create FAC Table Dataflow Refresh
# =================================================================================================================================================

# =================================================================================================================================================
# START -------- DIM Table Dataflow Refresh
# =================================================================================================================================================
foreach($Dataflow in $ListAllDataflows) {
    $DataflowId = $Dataflow.objectId
    $GetDataflowSchedulesApiCall = $PbiRestApi + "groups/" + $WorkspaceId + "/dataflows/" + $DataflowId + "/refreshes"
    $DataflowSchedule = Invoke-PowerBIRestMethod -Method GET -Url $GetDataflowSchedulesApiCall | ConvertFrom-Json
    $DataflowSchedule | Add-Member -MemberType  NoteProperty -Name 'Id_Dataflow'    -Value $Dataflow.objectId -Force
    $DataflowSchedule | Add-Member -MemberType  NoteProperty -Name 'ExtractionDate' -Value $DatePrefix -Force
    $DIM_Dataset_Refresh_DQ += $DataflowSchedule
}
# =================================================================================================================================================
# END -------- DIM Table Dataflow Refresh
# =================================================================================================================================================




# =================================================================================================================================================
# START -------- Datasts GET
# =================================================================================================================================================

$GetDatasetsApiCall = $PbiRestApi + "groups/" + $WorkspaceId + "/datasets"
$AllDatasets = Invoke-PowerBIRestMethod -Method GET -Url $GetDatasetsApiCall | ConvertFrom-Json
$ListAllDatasets_ = $AllDatasets.value
$ListAllDatasets = $ListAllDatasets_ | Where-Object {$_.name -notin @("Report Usage Metrics Model","Usage Metrics Report","Usage Metrics Report","Dashboard Usage Metrics Model")}
# =================================================================================================================================================
# END -------- Datasts GET
# =================================================================================================================================================

# =================================================================================================================================================
# START -------- Dataset DIMENSION TABLE
# =================================================================================================================================================
$ListAllDatasets | ForEach-Object{
    $_ | Add-Member -MemberType NoteProperty 'Id_Workspace'   -Value $WorkspaceId -Force
    $_ | Add-Member -MemberType NoteProperty 'WorkspaceName'  -Value $WorkspaceInfo.name -Force
    $_ | Add-Member -MemberType NoteProperty 'ExtractionDate' -Value $DatePrefix -Force
}
$DIM_Datasets += $ListAllDatasets
# =================================================================================================================================================
# END -------- Dataset DIMENSION TABLE
# =================================================================================================================================================



# =================================================================================================================================================
# START -------- Function Get Dataset Refresh 
# =================================================================================================================================================
Function GetDatasetRefreshResults {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true)][string]$DatasetId
    )
    $GetDatasetRefreshHistory = $PbiRestApi + "groups/" + $WorkspaceId + "/datasets/" + $DatasetId + "/refreshes"
    $DatasetRefreshHistory = Invoke-PowerBIRestMethod -Method GET -Url $GetDatasetRefreshHistory | ConvertFrom-Json
    return $DatasetRefreshHistory.value
}

# =================================================================================================================================================
# START -------- Function aplication to create FAC Table Dataset Refresh
# =================================================================================================================================================
foreach($Dataset in $ListAllDatasets) {
    $DatasetHistories = GetDatasetRefreshResults -DatasetId $Dataset.id
    foreach($DatasetHistory in $DatasetHistories) {
        Add-Member -InputObject $DatasetHistory -NotePropertyName 'Id_Dataset'     -NotePropertyValue $Dataset.id
        Add-Member -InputObject $DatasetHistory -NotePropertyName 'Id_Workspace'   -NotePropertyValue $WorkspaceId -Force
        Add-Member -InputObject $DatasetHistory -NotePropertyName 'WorkspaceName'  -NotePropertyValue $WorkspaceInfo.name -Force
        Add-Member -InputObject $DatasetHistory -NotePropertyName 'ExtractionDate' -NotePropertyValue $DatePrefix -Force

        $FAC_Dataset_Refresh += $DatasetHistory
    }
}
# =================================================================================================================================================
# END  -------- Function Get Dataset Refresh and Function aplication to create FAC Table Dataset Refresh
# =================================================================================================================================================
foreach($Dataset in $ListAllDatasets) {
    $Datasetid = $Dataset.id
    $GetUsersApiCall = $PbiRestApi + "groups/" + $WorkspaceId + "/datasets/" + $Datasetid + "/Users"
    $DatasetUs = Invoke-PowerBIRestMethod -Method GET -Url  $GetUsersApiCall
    $DatasetUs += $DatasetUsers  
} 

# =================================================================================================================================================
# START -------- DIM Table Dataset Refresh
# =================================================================================================================================================

foreach($Dataset in $ListAllDatasets) {
    $Datasetid = $Dataset.id
    $GetDatasetSchedulesApiCall = $PbiRestApi + "groups/" + $WorkspaceId + "/datasets/" + $Datasetid + "/refreshSchedule"
    $DatasetSchedule = Invoke-PowerBIRestMethod -Method GET -Url $GetDatasetSchedulesApiCall | ConvertFrom-Json
    Add-Member -InputObject $DatasetSchedule -NotePropertyName 'Id_Workspace'   -NotePropertyValue $WorkspaceId -Force
    Add-Member -InputObject $DatasetSchedule -NotePropertyName 'WorkspaceName'  -NotePropertyValue $WorkspaceInfo.name -Force
    Add-Member -InputObject $DatasetSchedule -NotePropertyName 'Id_Dataset'     -NotePropertyValue $Dataset.id
    Add-Member -InputObject $DatasetSchedule -NotePropertyName 'ExtractionDate' -NotePropertyValue $DatePrefix
    Add-Member -InputObject $DatasetSchedule -NotePropertyName 'DirectQuery'    -NotePropertyValue "False"
    Add-Member -InputObject $DatasetSchedule -NotePropertyName 'Frecuency'      -NotePropertyValue ""
    $DIM_Dataset_Refresh += $DatasetSchedule
}
# =================================================================================================================================================
# END -------- DIM Table Dataset Refresh
# =================================================================================================================================================

# =================================================================================================================================================
# START -------- DIM Table Dataset Refresh Direct Query
# =================================================================================================================================================
foreach($Dataset in $ListAllDatasets) {
    $DatasetID = $Dataset.id
    $DatasetScheduleDQ = Invoke-PowerBIRestMethod -Method GET -Url "datasets/$DatasetID/directQueryRefreshSchedule" | ConvertFrom-Json
    Add-Member -InputObject $DatasetScheduleDQ -NotePropertyName 'Id_Workspace'   -NotePropertyValue $WorkspaceId -Force
    Add-Member -InputObject $DatasetScheduleDQ -NotePropertyName 'WorkspaceName'  -NotePropertyValue $WorkspaceInfo.name -Force
    Add-Member -InputObject $DatasetScheduleDQ -NotePropertyName 'Id_Dataset'     -NotePropertyValue $Dataset.id
    Add-Member -InputObject $DatasetScheduleDQ -NotePropertyName 'ExtractionDate' -NotePropertyValue $DatePrefix
    Add-Member -InputObject $DatasetScheduleDQ -NotePropertyName 'DirectQuery'    -NotePropertyValue "True"
    Add-Member -InputObject $DatasetScheduleDQ -NotePropertyName 'enabled'        -NotePropertyValue ""
    $DIM_Dataset_Refresh_DQ += $DatasetScheduleDQ
}
}
# =================================================================================================================================================
# END -------- DIM Table Dataset Refresh Direct Query
# =================================================================================================================================================
$adminUsername = $credentials.UserName

foreach ($workspace in $workspaces) {
    $workspaceServer = $workspace.Name -replace " ", "%20" -replace "&", "%26" -replace "/", "%2F"
    $serverName = "$PbiConnectionAS$workspaceServer"
    $WorkspaceId = $workspace.Id

# =================================================================================================================================================
# START -------- Datasts GET
# =================================================================================================================================================
$GetDatasetsApiCall = $PbiRestApi + "groups/" + $WorkspaceId + "/datasets"
$AllDatasets = Invoke-PowerBIRestMethod -Method GET -Url $GetDatasetsApiCall | ConvertFrom-Json
$ListAllDatasets_ = $AllDatasets.value
$ListAllDatasets = $ListAllDatasets_ | Where-Object {$_.name -notin @("Report Usage Metrics Model","Usage Metrics Report","Usage Metrics Report","Dashboard Usage Metrics Model")}
# =================================================================================================================================================
# END -------- Datasts GET
# =================================================================================================================================================



# =================================================================================================================================================
# END -------- Datasts GET
# =================================================================================================================================================
    $databaseName_ = "Usage Metrics Report"
    $adminPassword = "Migrazione10!"
    #$adminUsername = "dealer.publisher@tmi-toyota.it"
    #$adminPassword = "Toyota2018!"
    $connectionString_CODEM = "Provider=MSOLAP;Data Source=$serverName;Initial Catalog=$databaseName_;User Id=$adminUsername;Password=$adminPassword;"
    $query2 = "evaluate
    SUMMARIZE('Report views',
    [AppName],[CreationTime],[DatasetName],[DistributionMethod],[ReportId],[UserKey],'Report rank'[ReportRank])"
    $result = Invoke-ASCmd -Server $serverName -ConnectionString $connectionString_CODEM -Query $query2
    $matches_ = @()
    $matches_ = [regex]::Matches($result, '<row xmlns="urn:schemas-microsoft-com:xml-analysis:rowset">(.*?)</row>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    foreach ($match in $matches_) {
        $matchesC1 = [regex]::Matches($match.Value, '<C1>(.*?)</C1>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC2 = [regex]::Matches($match.Value, '<C2>(.*?)</C2>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC3 = [regex]::Matches($match.Value, '<C3>(.*?)</C3>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC4 = [regex]::Matches($match.Value, '<C4>(.*?)</C4>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC5 = [regex]::Matches($match.Value, '<C5>(.*?)</C5>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC6 = [regex]::Matches($match.Value, '<C6>(.*?)</C6>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
            $propertyC1 = if ($matchesC1.Count -gt 0) {$matchesC1[0].Groups[1].Value}{""}
            $propertyC2 = if ($matchesC2.Count -gt 0) {$matchesC2[0].Groups[1].Value}{""}
            $propertyC3 = if ($matchesC3.Count -gt 0) {$matchesC3[0].Groups[1].Value}{""}
            $propertyC4 = if ($matchesC4.Count -gt 0) {$matchesC4[0].Groups[1].Value}{""}
            $propertyC5 = if ($matchesC5.Count -gt 0) {$matchesC5[0].Groups[1].Value}{""}
            $propertyC6 = if ($matchesC6.Count -gt 0) {$matchesC6[0].Groups[1].Value}{""}
        
            $entry = [PSCustomObject]@{
                Date                      = $propertyC1
                DataserName               = $propertyC2
                DistributionMethod        = $propertyC3
                Id_Report                 = $propertyC4
                User                      = $propertyC5
                Report_Rank               = $propertyC6 
                WorkspaceName             = $workspace.Name
                Id_Workspace              = $workspace.Id
            }
            $ReportViews += $entry
        }  
# =================================================================================================================================================
# START --------CODE M
# =================================================================================================================================================
foreach ($database in $ListAllDatasets) {
    $databaseName = $database.name
    $adminPassword = "Migrazione10!"
    #$adminUsername = "dealer.publisher@tmi-toyota.it"
    #$adminPassword = "Toyota2018!"

    # Construir la cadena de conexión
    $connectionString_CODEM = "Provider=MSOLAP;Data Source=$serverName;Initial Catalog=$databaseName;User Id=$adminUsername;Password=$adminPassword;"
    $SYSTEM ='$SYSTEM'
    $query2 = "Select [TableID],[Name],[QueryDefinition] from $System.TMSCHEMA_PARTITIONS Where [Type] = 4"
    $result = Invoke-ASCmd -Server $serverName -ConnectionString $connectionString_CODEM -Query $query2
    $matches_ = @()
    $matches_ = [regex]::Matches($result, '<row xmlns="urn:schemas-microsoft-com:xml-analysis:rowset">(.*?)</row>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    foreach ($match in $matches_) {
    $matchesC0 = [regex]::Matches($match.Value, '<C0>(.*?)</C0>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $matchesC1 = [regex]::Matches($match.Value, '<C1>(.*?)</C1>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $matchesC2 = [regex]::Matches($match.Value, '<C2>(.*?)</C2>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
        $propertyC0 = if ($matchesC0.Count -gt 0) {$matchesC0[0].Groups[1].Value}{""}
        $propertyC1 = if ($matchesC1.Count -gt 0) {$matchesC1[0].Groups[1].Value}{""}
        $propertyC2 = if ($matchesC2.Count -gt 0) {$matchesC2[0].Groups[1].Value}{""}
    
        $entry = [PSCustomObject]@{
            TableID                   = $propertyC0+"*"+$database.id
            TableName_cod             = $propertyC1
            M_Code                    = $propertyC2
            Id_Dataset                = $database.id
            DatasetName               = $database.name
            Id_Workspace              = $workspace.id
            Mode                      = "Partitions"
            ExtractionDate  = $DatePrefix
            }
        $extractCODEM += $entry
    }
}


foreach ($database in $ListAllDatasets) {
            $databaseName = $database.name
            $adminPassword = "Migrazione10!"
            #$adminUsername = "dealer.publisher@tmi-toyota.it"
            #$adminPassword = "Toyota2018!"
        
            # Construir la cadena de conexión
            $connectionString_CODEM = "Provider=MSOLAP;Data Source=$serverName;Initial Catalog=$databaseName;User Id=$adminUsername;Password=$adminPassword;"
            $SYSTEM ='$SYSTEM'
            $query2 = "Select * from $System.DISCOVER_M_EXPRESSIONS"
            $result = Invoke-ASCmd -Server $serverName -ConnectionString $connectionString_CODEM -Query $query2
            $matches_ = @()
            $matches_ = [regex]::Matches($result, '<row xmlns="urn:schemas-microsoft-com:xml-analysis:rowset">(.*?)</row>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            foreach ($match in $matches_) {
            $matchesC0 = [regex]::Matches($match.Value, '<C0>(.*?)</C0>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            $matchesC1 = [regex]::Matches($match.Value, '<C1>(.*?)</C1>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
                $propertyC0 = if ($matchesC0.Count -gt 0) {$matchesC0[0].Groups[1].Value}{""}
                $propertyC1 = if ($matchesC1.Count -gt 0) {$matchesC1[0].Groups[1].Value}{""}
            
                $entry = [PSCustomObject]@{
                    TableID                   = ""
                    TableName_cod             = $propertyC0
                    M_Code                    = $propertyC1
                    Id_Dataset                = $database.id
                    DatasetName               = $database.name
                    Id_Workspace              = $workspace.id
                    Mode                      = "Expression"
                    ExtractionDate  = $DatePrefix
        }
                $extractCODEM2 += $entry
    }
}


foreach ($database in $ListAllDatasets) {
            $databaseName = $database.name
            $adminPassword = "Migrazione10!"
            #$adminUsername = "dealer.publisher@tmi-toyota.it"
            #$adminPassword = "Toyota2018!"
        
            # Construir la cadena de conexión
            $connectionString_CODEM = "Provider=MSOLAP;Data Source=$serverName;Initial Catalog=$databaseName;User Id=$adminUsername;Password=$adminPassword;"
            $SYSTEM ='$SYSTEM'
            $query2 = "select * from $SYSTEM.TMSCHEMA_TABLES where [SystemFlags] = 0"
            $result = Invoke-ASCmd -Server $serverName -ConnectionString $connectionString_CODEM -Query $query2
            $matches_ = @()
            $matches_ = [regex]::Matches($result, '<row xmlns="urn:schemas-microsoft-com:xml-analysis:rowset">(.*?)</row>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            foreach ($match in $matches_) {
            $matchesC00 = [regex]::Matches($match.Value, '<C00>(.*?)</C00>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            $matchesC02 = [regex]::Matches($match.Value, '<C02>(.*?)</C02>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            $matchesC07 = [regex]::Matches($match.Value, '<C07>(.*?)</C07>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            $matchesC08 = [regex]::Matches($match.Value, '<C08>(.*?)</C08>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
                $propertyC00 = if ($matchesC00.Count -gt 0) {$matchesC00[0].Groups[1].Value}{""}
                $propertyC02 = if ($matchesC02.Count -gt 0) {$matchesC02[0].Groups[1].Value}{""}
                $propertyC07 = if ($matchesC07.Count -gt 0) {$matchesC07[0].Groups[1].Value}{""}
                $propertyC08 = if ($matchesC08.Count -gt 0) {$matchesC08[0].Groups[1].Value}{""}
            
                $entry = [PSCustomObject]@{
                    TableID                   = $propertyC00+"*"+$database.id
                    TableName                 = $propertyC02
                    ModifiedTime              = $propertyC07
                    StructureModifiedTime     = $propertyC08
                }
                $extractCODEM3 += $entry
            }
            
        }
    

foreach ($database in $ListAllDatasets) {
        $databaseName = $database.name
        $adminPassword = "Migrazione10!"
        #$adminUsername = "dealer.publisher@tmi-toyota.it"
        #$adminPassword = "Toyota2018!"
    
        $connectionString_CODEM = "Provider=MSOLAP;Data Source=$serverName;Initial Catalog=$databaseName;User Id=$adminUsername;Password=$adminPassword;"
        $SYSTEM ='$SYSTEM'
        $query2 = "select * from $SYSTEM.TMSCHEMA_MEASURES"
        $result = Invoke-ASCmd -Server $serverName -ConnectionString $connectionString_CODEM -Query $query2
        $matches_ = @()
        $matches_ = [regex]::Matches($result, '<row xmlns="urn:schemas-microsoft-com:xml-analysis:rowset">(.*?)</row>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        foreach ($match in $matches_) {
        $matchesC01 = [regex]::Matches($match.Value, '<C01>(.*?)</C01>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC02 = [regex]::Matches($match.Value, '<C02>(.*?)</C02>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC04 = [regex]::Matches($match.Value, '<C04>(.*?)</C04>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC05 = [regex]::Matches($match.Value, '<C05>(.*?)</C05>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC09 = [regex]::Matches($match.Value, '<C09>(.*?)</C09>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC10 = [regex]::Matches($match.Value, '<C10>(.*?)</C10>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $matchesC13 = [regex]::Matches($match.Value, '<C13>(.*?)</C13>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
            $propertyC01 = if ($matchesC01.Count -gt 0) {$matchesC01[0].Groups[1].Value}{""}
            $propertyC02 = if ($matchesC02.Count -gt 0) {$matchesC02[0].Groups[1].Value}{""}
            $propertyC04 = if ($matchesC04.Count -gt 0) {$matchesC04[0].Groups[1].Value}{""}
            $propertyC05 = if ($matchesC05.Count -gt 0) {$matchesC05[0].Groups[1].Value}{""}
            $propertyC09 = if ($matchesC09.Count -gt 0) {$matchesC09[0].Groups[1].Value}{""}
            $propertyC10 = if ($matchesC10.Count -gt 0) {$matchesC10[0].Groups[1].Value}{""}
            $propertyC13 = if ($matchesC13.Count -gt 0) {$matchesC13[0].Groups[1].Value}{""}
        
            $entry = [PSCustomObject]@{
                Id_Dataset               = $database.id
                TableID                  = $propertyC01+"*"+$database.id
                MeasureName              = $propertyC02
                MeasureType              = $propertyC04
                Expression               = $propertyC05
                ModifiedTime             = $propertyC09
                StructureModifiedTime    = $propertyC10
                ErrorMessage             = $propertyC13
                ExtractionDate           = $DatePrefix
            }
            $DatasetMeasure += $entry
        }   
    }
}


# =================================================================================================================================================
# END --------CODE M
# =================================================================================================================================================


# =================================================================================================================================================
# START -------- EXTRACT THE SQL TABLES AND THEN LOOK THEM UP IN THE M CODE
# =================================================================================================================================================
$tabelleSQL = @()
$columnSQL = @()

$server_1   = $config.Server_1
$Database_1 = $config.Database_1
$UserName_1 = $config.UserName_1
$Password_1 = $config.Password_1
if ($config.Authentication_1 -eq "Azure Active Directory")
{$Authentication_1 = ";Authentication=Active Directory Password"}
else {$Authentication_1 = ""}

 
$connectionString_1 = "Server=$Server_1;Database=$Database_1;User Id=$UserName_1;Password=$Password_1$Authentication_1"
$connection_1 = New-Object System.Data.SqlClient.SqlConnection
$connection_1.ConnectionString = $connectionString_1
$connection_1.Open()

$sqlQuery_1 = 'SELECT table_name FROM information_schema.tables'
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
$command_1 = Invoke-Sqlcmd -ServerInstance $Server_1 -Database $Database_1  -Username $UserName_1 -Password $Password_1 -Query   $sqlQuery_1 -AccessToken $AccessToken
foreach ($row in $command_1) {
    $Value = $row["table_name"] 
    $tabelleSQL += $Value}

try {
    $sqlQuery_1 = 'SELECT table_name FROM information_schema.tables'
    $command_1 = $connection_1.CreateCommand()
    $command_1.CommandText = $sqlQuery_1
    $reader_1 = $command_1.ExecuteReader()
 
    while ($reader_1.Read()) {
        $tableName = $reader_1['table_name']
        $tabelleSQL += $tableName
    }
}
finally {
    $connection_1.Close()
}
$tabelleSQL

$server_2   = $config.Server_2
$Database_2 = $config.Database_2
$UserName_2 = $config.UserName_2
$Password_2 = $config.Password_2
if ($config.Authentication_2 -eq "Azure Active Directory")
{$Authentication_2 = ";Authentication=Active Directory Password"}
else {$Authentication_2 = ""}


$connectionString_2 = "Server=$Server_2;Database=$Database_2;User Id=$UserName_2;Password=$Password_2$Authentication_2"
$connection_2 = New-Object System.Data.SqlClient.SqlConnection
$connection_2.ConnectionString = $connectionString_2
$connection_2.Open()

$sqlQuery_2 = 'SELECT table_name FROM information_schema.tables'
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
$command_2 = Invoke-Sqlcmd -ServerInstance $Server_2 -Database $Database_2  -Username $UserName_2 -Password $Password_2 -Query   $sqlQuery_2 -AccessToken $AccessToken
foreach ($row in $command_2) {
    $Value = $row["table_name"] 
    $tabelleSQL += $Value}

try {
    $sqlQuery_2 = 'SELECT table_name FROM information_schema.tables'
    $command_2 = $connection_2.CreateCommand()
    $command_2.CommandText = $sqlQuery_2
    $reader_2 = $command_2.ExecuteReader()

    while ($reader_2.Read()) {
        $tableName = $reader_2['table_name']
        $tabelleSQL += $tableName
    }
}
finally {
    $connection_2.Close()
}
# =================================================================================================================================================
# END -------- EXTRACT THE SQL TABLES AND THEN LOOK THEM UP IN THE M CODE
# =================================================================================================================================================


# =================================================================================================================================================
# START -------- EXTRACT THE SQL COLUMNS AND THEN LOOK THEM UP IN THE M CODE
# =================================================================================================================================================


$connectionString_1 = "Server=$Server_1;Database=$Database_1;User Id=$UserName_1;Password=$Password_1$Authentication_1"
$connection_1 = New-Object System.Data.SqlClient.SqlConnection
$connection_1.ConnectionString = $connectionString_1
$connection_1.Open()

$sqlQuery_1 = 'select distinct column_name from INFORMATION_SCHEMA.columns'
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
$command_1 = Invoke-Sqlcmd -ServerInstance $Server_1 -Database $Database_1  -Username $UserName_1 -Password $Password_1 -Query   $sqlQuery_1 -AccessToken $AccessToken
foreach ($row in $command_1) {
    $Value = $row["column_name"] 
    $columnSQL += $Value}

try {
    $sqlQuery_1 = 'select distinct column_name from INFORMATION_SCHEMA.columns'
    $command_1 = $connection_1.CreateCommand()
    $command_1.CommandText = $sqlQuery_1
    $reader_1 = $command_1.ExecuteReader()
 
    while ($reader_1.Read()) {
        $columnName = $reader_1['column_name']
        $columnSQL += $columnName
    }
}
finally {
    $connection_1.Close()
}

$connectionString_2 = "Server=$Server_2;Database=$Database_2;User Id=$UserName_2;Password=$Password_2$Authentication_2"
$connection_2 = New-Object System.Data.SqlClient.SqlConnection
$connection_2.ConnectionString = $connectionString_2
$connection_2.Open()

$sqlQuery_2 = 'select distinct column_name from INFORMATION_SCHEMA.columns'
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
$command_2 = Invoke-Sqlcmd -ServerInstance $Server_2 -Database $Database_2  -Username $UserName_2 -Password $Password_2 -Query   $sqlQuery_2 -AccessToken $AccessToken
foreach ($row in $command_2) {
    $Value = $row["column_name"] 
    $columnSQL += $Value}

try {
    $sqlQuery_2 = 'select distinct column_name from INFORMATION_SCHEMA.columns'
    $command_2 = $connection_2.CreateCommand()
    $command_2.CommandText = $sqlQuery_2
    $reader_2 = $command_2.ExecuteReader()
 
    while ($reader_2.Read()) {
        $columnName = $reader_2['column_name']
        $columnSQL += $columnName
    }
}
finally {
    $connection_2.Close()
}
# =================================================================================================================================================
# END -------- EXTRACT THE SQL TABLES AND THEN LOOK THEM UP IN THE M CODE
# =================================================================================================================================================

$codeMFinale3 = $extractCODEM + $extractCODEM2
$codeMFinale2 = $codeMFinale3 | Where-Object {$_.M_Code -ne $null} 
$codeMFinale1 = $codeMFinale2 | Where-Object {$_.TableName_cod -notlike "Errors in*" -and $_.TableName_cod -notlike "Errori in*" -and $_.TableName_cod -notlike "File di esempio*" -and $_.TableName_cod -notlike "Transform File*" -and $_.TableName_cod -notlike "DateTableTemplate*" -and $_.TableName_cod -notlike "LocalDateTable*" -and $_.TableName_cod -notlike "Trasforma file*" -and $_.TableName_cod -notlike "Parameter*" -and $_.TableName_cod -notlike "Parametro*" -and $_.TableName_cod -notlike "Transform Sample File*" -and $_.TableName_cod -notlike "Sample File*" } 


$ID = 1
foreach($obj in $codeMFinale1){
    $obj | Add-Member -MemberType NoteProperty -Name "ID" -Value $ID -Force
    $ID++
}
foreach ($obj1Item in $codeMfinale1) {
    foreach ($obj2Item in $extractCODEM3) {
        if ($obj1Item.TableID -eq $obj2Item.TableID) {
            $obj1Item | Add-Member -MemberType NoteProperty -Name "TableName"                 -Value $obj2Item.TableName                   -Force
            $obj1Item | Add-Member -MemberType NoteProperty -Name "ModifiedTime"              -Value $obj2Item.ModifiedTime                -Force
            $obj1Item | Add-Member -MemberType NoteProperty -Name "StructureModifiedTime"     -Value $obj2Item.StructureModifiedTime       -Force
        }
    }
}

$codeMFinale = $codeMFinale1 | Where-Object {$_.TableName -notlike "Errors in*" -and $_.TableName -notlike "Errori in*" -and $_.TableName -notlike "File di esempio*" -and $_.TableName -notlike "Transform File*" -and $_.TableName -notlike "DateTableTemplate*" -and $_.TableName -notlike "LocalDateTable*" -and $_.TableName -notlike "Trasforma file*" -and $_.TableName -notlike "Parameter*" -and $_.TableName -notlike "Transform Sample File*" -and $_.TableName -notlike "Sample File*" } 
$CODEM_json = $codeMFinale | ConvertTo-Json -Depth 2
$CODEMJson = ConvertFrom-Json $CODEM_json

# =================================================================================================================================================
# START -------- FIND SQL TABLES AND COLUMNS IN M CODE
# =================================================================================================================================================

$MatchTablesColumns = @()
$tabelleSQL = $tabelleSQL -replace  "[\[\]]", "''"

$ServerData = 'syn-tmi-prod.sql.azuresynapse.net','sql-tmidata-prod.database.windows.net'
$Source = 'Sql.Database', 'Excel.Workbook', 'SharePoint.Files','Table.FromRows', 'Odbc.DataSource','PowerBI.Dataflows', 'AnalysisServices.Database', 'UsageMetricsDataConnector','Folder.Files','Table.NestedJoin','Table.Combine','Table.FromList','Csv.Document','DateTime.LocalNow', 'SharePoint.Tables', 'PowerPlatform.Dataflows'
$database_sql = 'tmi_dwh','syndp_tmi_dwi_realtime_datamart'


function TablesSearch($jsonObject, $array) {
    $result = @()
    foreach ($value in $array) {
        $regex = [regex]::new($value,'IgnoreCase')
        $matches_ = $regex.Matches($jsonObject.M_Code.ToString())
        if ($matches_.Count -gt 0) {
            foreach ($match in $matches_) {
                $matchValue = $match.Value
                if ($matchValue -notin $result) {
                    $result += $matchValue
                }
            }
        }
    }
    return $result
}


foreach ($jsonObject in $codeMFinale) {
    if($null -ne $jsonObject.M_Code){
    $id          = $jsonObject.ID
    $matchserver = TablesSearch $jsonObject  $ServerData
    $matchSource = TablesSearch $jsonObject  $Source | Select-Object -First 1
    $matchDatabase = TablesSearch $jsonObject  $database_sql Select-Object -First 1
    $matchTables =  if ($matchSource -eq 'Sql.Database') { TablesSearch $jsonObject  $tabelleSQL }
    $matchColumns = if ($matchSource -eq 'Sql.Database') { TablesSearch $jsonObject  $columnSQL }


    $concatenatedValues = $matchTables -join ','
    $uniqueColumns = $matchColumns | Get-Unique
    $concatenatedValues_col = $uniqueColumns -join ','
    $jsonResult = [PSCustomObject]@{
        ID = $jsonObject.ID
        Tables_SQL = $concatenatedValues
        Server = $matchserver
        Source = $matchSource
        Database_SQL = $matchDatabase
        Columns_SQL = $concatenatedValues_col
        }
    $MatchTablesColumns += $jsonResult
    }
}

# =================================================================================================================================================
# END --------CONVERT JSON
# =================================================================================================================================================





# =================================================================================================================================================
# START --------CONVERT JSON
# =================================================================================================================================================

$DIM_Reports_Json = $DIM_Reports | ConvertTo-Json -Depth 2
$DIM_ReportsJson = ConvertFrom-Json $DIM_Reports_Json

$DatiReports_Json = $DatiReports | ConvertTo-Json -Depth 2
$DatiReportsJson = ConvertFrom-Json $DatiReports_Json

$DIM_Datasets_Json = $DIM_Datasets | ConvertTo-Json -Depth 2
$DIM_DatasetsJson = ConvertFrom-Json $DIM_Datasets_Json

$DIM_Dataset_Refresh_DIM_Json = $DIM_Dataset_Refresh | ConvertTo-Json -Depth 2
$DIM_Dataset_RefreshDIMJson = ConvertFrom-Json $DIM_Dataset_Refresh_DIM_Json

$DIM_Dataflows_Json = $DIM_Dataflows | ConvertTo-Json -Depth 2
$DIM_DataflowsJson = ConvertFrom-Json $DIM_Dataflows_Json

$FAC_Dataflow_Refresh_Json = $FAC_Dataflow_Refresh | ConvertTo-Json -Depth 2
$FAC_Dataflow_RefreshJson = ConvertFrom-Json $FAC_Dataflow_Refresh_Json

$FAC_Dataset_Refresh_Json = $FAC_Dataset_Refresh  | ConvertTo-Json -Depth 2
$FAC_Dataset_RefreshJson = ConvertFrom-Json $FAC_Dataset_Refresh_Json

$DIM_Dataset_Refresh_DQ_Json = $DIM_Dataset_Refresh_DQ | ConvertTo-Json -Depth 2
$DIM_Dataset_Refresh_DQJson = ConvertFrom-Json $DIM_Dataset_Refresh_DQ_Json

$DIM_Workspaces_Json = $DIM_Workspaces | ConvertTo-Json -Depth 2
$DIM_WorkspacesJson = ConvertFrom-Json $DIM_Workspaces_Json

$FAC_Database_Measure_Json = $DatasetMeasure | ConvertTo-Json -Depth 2
$FAC_Database_MeasureJson = ConvertFrom-Json $FAC_Database_Measure_Json

foreach ($obj1Item in $codeMfinale) {
    foreach ($obj2Item in $MatchTablesColumns) {
        if ($obj1Item.ID -eq $obj2Item.ID) {
            $obj1Item | Add-Member -MemberType NoteProperty -Name "Tables_SQL"   -Value $obj2Item.Tables_SQL   -Force
            $obj1Item | Add-Member -MemberType NoteProperty -Name "Server"       -Value $obj2Item.Server       -Force
            $obj1Item | Add-Member -MemberType NoteProperty -Name "Source"       -Value $obj2Item.Source       -Force
            $obj1Item | Add-Member -MemberType NoteProperty -Name "Database_SQL" -Value $obj2Item.Database_SQL -Force
            $obj1Item | Add-Member -MemberType NoteProperty -Name "Columns_SQL"  -Value $obj2Item.Columns_SQL  -Force
        }
    }
}
$CODEM_json = $codeMfinale | ConvertTo-Json -Depth 2
$CODEMJson = ConvertFrom-Json $CODEM_json

$FAC_ReportViews_json = $ReportViews | ConvertTo-Json -Depth 2
$FAC_ReportViewsJson = ConvertFrom-Json $FAC_ReportViews_json

# =================================================================================================================================================
# END --------CONVERT JSON
# =================================================================================================================================================




# =================================================================================================================================================
# START ----------- TABLES TRUNCATE 
# =================================================================================================================================================


# =================================================================================================================================================
# START ----------- FORM CONECTION SQL
# =================================================================================================================================================
$Server = $config.Server
$Database = $config.Database
$UserName= $config.UserName
$Password = $config.Password
# =================================================================================================================================================
# END ----------- FORM CONECTION SQL
# =================================================================================================================================================



function truncateSQL {
    param (
        [string] $table
    ) 
        $TRUNCATE = "TRUNCATE TABLE $table"
        $command = $connection.CreateCommand()
        $command.CommandText = $TRUNCATE
        $command.ExecuteNonQuery() 
}

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDimensionReports"   -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDimensionDataflows"   -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDimensionWorkspaces"   -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDimensionDatasets"  -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDimensionDatasetRefresh" -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDimensionDataSource"  -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableFactRefreshDataset" -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableFactRefreshDataflow" -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableFactCODEM" -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableDatabaseMeasure" -AccessToken $AccessToken

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query "TRUNCATE TABLE $tableReportViews" -AccessToken $AccessToken

try {
    $connectionString = "Server=$Server;Database=$Database;User Id=$UserName;Password=$Password$Authentication"
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString
    $connection.Open()
    Write-Host $($connection.State)
} 
catch {
    Write-Host $_.Exception.Message
    Write-Host $($_.Exception.StackTrace)
}

    # Crear la tabla dinámicamente
    truncateSQL -table $tableDimensionReports 
    truncateSQL -table $tableDimensionDataflows 
    truncateSQL -table $tableDimensionWorkspaces 
    truncateSQL -table $tableDimensionDatasets
    truncateSQL -table $tableDimensionDatasetRefresh 
    truncateSQL -table $tableDimensionDataSource 
    truncateSQL -table $tableFactRefreshDataset
    truncateSQL -table $tableFactRefreshDataflow 
    truncateSQL -table $tableFactCODEM
    truncateSQL -table $tableDatabaseMeasure
    truncateSQL -table $tableReportViews

# =================================================================================================================================================
# END ----------- TABLES TRUNCATE 
# =================================================================================================================================================

# =================================================================================================================================================
# INSERT INTO -------- DIM Table REPORTS
# =================================================================================================================================================
foreach($obj in $DIM_ReportsJson) {
    $Id_Report      = $obj.Id_Report        -replace "'", "''"
    $DatasetName    = $obj.DatasetName      -replace "'", "''"
    $Id_Dataset     = $obj.Id_Dataset       -replace "'", "''"
    $Owner          = $obj.Owner            -replace "'", "''"
    $ReportName     = $obj.ReportName       -replace "'", "''"
    $ReportWebUrl   = $obj.ReportWebUrl     -replace "'", "''"
    $Id_Workspace   = $obj.Id_Workspace     -replace "'", "''"
    $WorkspaceName  = $obj.WorkspaceName    -replace "'", "''"
    $ExtractionDate = $obj.ExtractionDate   -replace "'", "''"
    
$query_Insert_Dim_Reports = "INSERT INTO $tableDimensionReports
(
	[Id_Report],
	[DatasetName],
	[Id_Dataset],
	[Owner],
	[ReportName],
	[ReportWebUrl],
	[Id_Workspace],
	[WorkspaceName], 
    [ExtractionDate] 
)VALUES(
    '$Id_Report',
    '$DatasetName',
    '$Id_Dataset',
    '$Owner ',
    '$ReportName',
    '$ReportWebUrl',
    '$Id_Workspace',
    '$WorkspaceName',
    '$ExtractionDate'

)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
    {   $command = $connection.CreateCommand()
        $command.CommandText = $query_Insert_Dim_Reports
        $command.ExecuteNonQuery()
            }
        else {
        $AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
        Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_Insert_Dim_Reports -AccessToken $AccessToken
}

}



# =================================================================================================================================================
# INSERT INTO -------- DIM Table DATAFLOWS
# =================================================================================================================================================
foreach($obj in $DIM_DataflowsJson) {
    $Id_Dataflow         = $obj.objectId        -replace "'", "''"
	$DataflowName        = $obj.name            -replace "'", "''"
	$DataflowDescription = $obj.description     -replace "'", "''"
	$ConfiguredBy        = $obj.configuredBy    -replace "'", "''"
	$Users               = $obj.users           -replace "'", "''"
	$Id_Workspace        = $obj.Id_Workspace    -replace "'", "''"
	$WorkspaceName       = $obj.WorkspaceName   -replace "'", "''"
	$ExtractionDate      = $obj.ExtractionDate  -replace "'", "''"

$query_Insert_Dim_Dataflows = "INSERT INTO $tableDimensionDataflows
(
	[Id_Dataflow],
	[DataflowName],
	[DataflowDescription],
	[ConfiguredBy],
	[Users],
	[Id_Workspace],
	[WorkspaceName],
	[ExtractionDate] 
)VALUES(
    '$Id_Dataflow',
	'$DataflowName',
	'$DataflowDescription',
	'$ConfiguredBy',
	'$Users',
	'$Id_Workspace',
	'$WorkspaceName',
	'$ExtractionDate' 
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")

        {$command = $connection.CreateCommand()
        $command.CommandText = $query_Insert_Dim_Dataflows
        $command.ExecuteNonQuery()}
else{
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_Insert_Dim_Dataflows -AccessToken $AccessToken}

}



# =================================================================================================================================================
# INSERT INTO -------- DIM Table WORKSPACES
# =================================================================================================================================================
foreach($obj in $DIM_WorkspacesJson) {
$Id_Workspace                = $obj.id                           -replace "'", "''"
$WorkspaceName               = $obj.name                         -replace "'", "''"
$IsReadOnly                  = $obj.isReadOnly                   -replace "'", "''"
$isOnDedicatedCapacity       = $obj.isOnDedicatedCapacity        -replace "'", "''"
$CapacityId                  = $obj.capacityId                   -replace "'", "''"
$defaultDatasetStorageFormat = $obj.defaultDatasetStorageFormat  -replace "'", "''"
$ExtractionDate              = $obj.ExtractionDate               -replace "'", "''"

$query_insert_Dim_workspaces = "INSERT INTO $tableDimensionWorkspaces
(
	[Id_Workspace],
	[WorkspaceName],
	[IsReadOnly],
	[isOnDedicatedCapacity],
	[CapacityId],
	[defaultDatasetStorageFormat],
    [ExtractionDate] 
)VALUES(
    '$Id_Workspace',
    '$WorkspaceName',
    '$IsReadOnly',
    '$isOnDedicatedCapacity',
    '$CapacityId',
    '$defaultDatasetStorageFormat',
    '$ExtractionDate' 
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_Dim_workspaces
$command.ExecuteNonQuery()}
else {

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_insert_Dim_workspaces -AccessToken $AccessToken
}
}



# =================================================================================================================================================
# INSERT INTO -------- DIM Table DATASETS
# =================================================================================================================================================
foreach($obj in $DIM_DatasetsJson) {
    $Id_Dataset               = $obj.id                        -replace "'", "''"
    $Dataset_Name             = $obj.name                      -replace "'", "''"
    $WebUrl                   = $obj.webUrl                    -replace "'", "''"
    $AddRowsAPIEnabled        = $obj.addRowsAPIEnabled         -replace "'", "''"
    $ConfiguredBy             = $obj.configuredBy              -replace "'", "''"
    $isOnPremGatewayRequired  = $obj.isOnPremGatewayRequired   -replace "'", "''"
    $TargetStorageMode        = $obj.targetStorageMode         -replace "'", "''"
    $CreatedDate              = $obj.createdDate               -replace "'", "''"
    $Id_Workspace             = $obj.Id_Workspace              -replace "'", "''"
    $WorkspaceName            = $obj.WorkspaceName             -replace "'", "''"
    $ExtractionDate           = $obj.ExtractionDate

$query_insert_Dim_datasets = "INSERT INTO $tableDimensionDatasets
(
	[Id_Dataset],
	[Dataset_Name],
	[WebUrl],
	[AddRowsAPIEnabled],
	[ConfiguredBy],
	[isOnPremGatewayRequired],
	[TargetStorageMode],
	[CreatedDate],
	[Id_Workspace],
	[WorkspaceName],
	[ExtractionDate] 
)VALUES(
    '$Id_Dataset',
    '$Dataset_Name',
    '$WebUrl',
    '$AddRowsAPIEnabled',
    '$ConfiguredBy',
    '$isOnPremGatewayRequired',
    '$TargetStorageMode',
    '$CreatedDate',
    '$Id_Workspace',
    '$WorkspaceName',
    '$ExtractionDate' 
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_Dim_datasets
$command.ExecuteNonQuery()}

else {
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_insert_Dim_datasets -AccessToken $AccessToken
}
}




# =================================================================================================================================================
# INSERT INTO -------- DIM Table DATASETS REFRESH
# =================================================================================================================================================
foreach($obj in $DIM_Dataset_RefreshDIMJson) {
    $Id_Dataset      = $obj.Id_Dataset       -replace "'", "''"
    $DirectQuery     = $obj.DirectQuery      -replace "'", "''"
    $LocalTimeZoneId = $obj.localTimeZoneId  -replace "'", "''"
    $Enable          = $obj.enabled          -replace "'", "''"
    $Times           = $obj.times            -replace "'", "''"
    $Days            = $obj.days             -replace "'", "''"
    $Frecuency       = $obj.Frecuency        -replace "'", "''"
    $Id_Workspace    = $obj.Id_Workspace     -replace "'", "''"
    $ExtractionDate  = $obj.ExtractionDate

    
$query_insert_Dim_datasets_Refresh = "INSERT INTO $tableDimensionDatasetRefresh
(
	[Id_Dataset],
	[DirectQuery],
	[LocalTimeZoneId],
	[Enable],
	[Times],
	[Days],
	[Frecuency],
	[Id_Workspace],
	[ExtractionDate] 
)VALUES(
    '$Id_Dataset',
    '$DirectQuery',
    '$LocalTimeZoneId',
    '$Enable',
    '$Times',
    '$Days',
    '$Frecuency',
    '$Id_Workspace',
    '$ExtractionDate' 

)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_Dim_datasets_Refresh
$command.ExecuteNonQuery()}
else {
    $AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
    Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query  $query_insert_Dim_datasets_Refresh -AccessToken $AccessToken     
}

}


# =================================================================================================================================================
# INSERT INTO -------- DIM Table DATASETS REFRESH DQ
# =================================================================================================================================================
foreach($obj in $DIM_Dataset_Refresh_DQJson) {
    $Id_Dataset      = $obj.Id_Dataset       -replace "'", "''"
    $DirectQuery     = $obj.DirectQuery      -replace "'", "''"
    $LocalTimeZoneId = $obj.localTimeZoneId  -replace "'", "''"
    $Enable          = $obj.enabled          -replace "'", "''"
    $Times           = $obj.times            -replace "'", "''"
    $Days            = $obj.days             -replace "'", "''"
    $Frecuency       = $obj.frequency        -replace "'", "''"
    $Id_Workspace    = $obj.Id_Workspace     -replace "'", "''"
    $ExtractionDate  = $obj.ExtractionDate

    
$query_insert_Dim_datasets_Refresh = "INSERT INTO $tableDimensionDatasetRefresh
(
	[Id_Dataset],
	[DirectQuery],
	[LocalTimeZoneId],
	[Enable],
	[Times],
	[Days],
	[Frecuency],
	[Id_Workspace],
	[ExtractionDate] 
)VALUES(
    '$Id_Dataset',
    '$DirectQuery',
    '$LocalTimeZoneId',
    '$Enable',
    '$Times',
    '$Days',
    '$Frecuency',
    '$Id_Workspace',
    '$ExtractionDate' 

)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_Dim_datasets_Refresh
$command.ExecuteNonQuery()
}
else {   
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query  $query_insert_Dim_datasets_Refresh -AccessToken $AccessToken
}
}


# =================================================================================================================================================
# INSERT INTO -------- DIM Table DATA SOURCE
# =================================================================================================================================================
foreach($obj in $DatiReportsJson) {
    $Id_Dataset      = $obj.Id_Dataset        -replace "'", "''"
    $DatasetName     = $obj.Datasetname       -replace "'", "''"
    $DatasourceType  = $obj.DatasourceType    -replace "'", "''"
    $Details         = $obj.Details           -replace "'", "''"
    $Refresh         = $obj.Refresh           -replace "'", "''"
    $Id_Report       = $obj.Id_Report         -replace "'", "''"
    $ReportName      = $obj.ReportName        -replace "'", "''"
    $Id_Workspace    = $obj.Id_Workspace      -replace "'", "''"
    $WorkspaceName   = $obj.WorkspaceName     -replace "'", "''"
    $ExtractionDate  = $obj.ExtractionDate 
    
$query_insert_Dim_dataSource = "INSERT INTO $tableDimensionDataSource
(
	[Id_Dataset],
	[DatasetName],
	[DatasourceType],
	[Details],
	[Refresh],
	[Id_Report],
	[ReportName],
	[Id_Workspace],
	[WorkspaceName],
    [ExtractionDate]  
)VALUES(
    '$Id_Dataset',
	'$DatasetName',
	'$DatasourceType',
	'$Details',
	'$Refresh',
	'$Id_Report',
	'$ReportName',
	'$Id_Workspace',
	'$WorkspaceName',
    '$ExtractionDate' 
 

)"

if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_Dim_dataSource
$command.ExecuteNonQuery()
}
else {
    
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query   $query_insert_Dim_dataSource -AccessToken $AccessToken
}
}




# =================================================================================================================================================
# INSERT INTO -------- FAC Table DATASET REFRESH
# =================================================================================================================================================
foreach($obj in $FAC_Dataset_RefreshJson) {

$Transaccion_Id          = $obj.id                    -replace "'", "''"
$Id_Dataset              = $obj.Id_Dataset            -replace "'", "''"
$RefreshType             = $obj.refreshType           -replace "'", "''"
$StartTime               = $obj.startTime             -replace "'", "''"
$EndTime                 = $obj.endTime               -replace "'", "''"
$Status                  = $obj.status                -replace "'", "''"
$Id_Workspace            = $obj.Id_Workspace          -replace "'", "''"
$WorkspaceName           = $obj.WorkspaceName         -replace "'", "''"
$ExtractionDate          = $obj.ExtractionDate        -replace "'", "''"
$serviceExceptionJson    = $obj.serviceExceptionJson  -replace "'", "''"
$refreshAttempts         = $obj.refreshAttempts       -replace "'", "''"

$query_insert_RefreshDataset = "INSERT INTO $tableFactRefreshDataset
(
	[Transaccion_Id],
	[Id_Dataset],
	[RefreshType],
	[StartTime],
	[EndTime],
	[Status],
	[Id_Workspace],
	[WorkspaceName],
	[ExtractionDate],
	[serviceExceptionJson],
	[refreshAttempts] 
)VALUES(
    '$Transaccion_Id',
    '$Id_Dataset',
    '$RefreshType',
    '$StartTime',
    '$EndTime',
    '$Status',
    '$Id_Workspace',
    '$WorkspaceName',
    '$ExtractionDate',
    '$serviceExceptionJson',
    '$refreshAttempts' 
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_RefreshDataset
$command.ExecuteNonQuery()
}
else {
    
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query  $query_insert_RefreshDataset -AccessToken $AccessToken
}
}

# =================================================================================================================================================
# INSERT INTO -------- FAC Table DATAFLOWS REFRESH
# =================================================================================================================================================
foreach($obj in $FAC_Dataflow_RefreshJson) {
    $Transaccion_Id  = $obj.id                   -replace "'", "''"
    $Id_Dataflow     = $obj.Id_Dataflow          -replace "'", "''"
    $RefreshType     = $obj.refreshType          -replace "'", "''"
    $StartTime       = $obj.startTime            -replace "'", "''"
    $EndTime         = $obj.endTime              -replace "'", "''"
    $Status          = $obj.status               -replace "'", "''"
    $Id_Workspace    = $obj.Id_Workspace         -replace "'", "''"
    $WorkspaceName   = $obj.WorkspaceName        -replace "'", "''"
    $ExtractionDate  = $obj.ExtractionDate

    $query_insert_RefreshDataflow = "INSERT INTO $tableFactRefreshDataflow
(
	[Transaccion_Id],
	[Id_Dataflow],
	[RefreshType],
	[StartTime],
	[EndTime],
	[Status],
	[Id_Workspace],
	[WorkspaceName],
	[ExtractionDate] 
)VALUES(
    '$Transaccion_Id',
    '$Id_Dataflow',
    '$RefreshType', 
    '$StartTime',
    '$EndTime',
    '$Status',
    '$Id_Workspace',
    '$WorkspaceName',
    '$ExtractionDate' 
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_RefreshDataflow
$command.ExecuteNonQuery()
}
else {
    

$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_insert_RefreshDataflow -AccessToken $AccessToken
}
}

# =================================================================================================================================================
# INSERT INTO -------- FAC Table M CODE
# =================================================================================================================================================
foreach($obj in $CODEMJson) {
    $ID = $obj.ID
    $Id_Dataset      = $obj.Id_Dataset           -replace "'", "''"
    $DatasetName     = $obj.DatasetName          -replace "'", "''"
    $Id_Workspace    = $obj.Id_Workspace         -replace "'", "''"
    $MCODE           = $obj.M_Code               -replace "'", "''" 
    $LENGHT          = if ($MCODE.Length -le 7990){$MCODE.Length}else {7990}
    $MCODETRUN       = $MCODE.Substring(0, $LENGHT)    
    $LoadToReport    = $obj.Mode                  -replace "'", "''"
    $TableName_cod   = $obj.TableName_cod        -replace "'", "''"
    $Table_Source    = $obj.Source               -replace "'", "''"
    $SQL_Server      = $obj.Server               -replace "'", "''"
	$SQL_Database    = $obj.Database_SQL         -replace "'", "''"
	$TableName_SQL   = $obj.Tables_SQL           -replace "'", "''"
	$Columns_SQL     = $obj.Columns_SQL          -replace "'", "''"
    $ExtractionDate  = $obj.ExtractionDate
    $TableName       = $obj.TableName             -replace "'", "''"
	$ModifiedTime    = $obj.ModifiedTime
    $StructureModifiedTime = $obj.StructureModifiedTime

    $query_insert_CodeM = "INSERT INTO $tableFactCODEM
(   [ID],
    [Id_Dataset],
    [DatasetName],
    [Id_Workspace],
	[M_Code],
	[LoadToReport],
	[TableName(PBI)],
	[Table_Source],
    [SQL_Server],
	[SQL_Database],
	[TableName(SQL)],
	[Columns(SQL)],
    [ExtractionDate],
    [TableName],
	[ModifiedTime],
    [StructureModifiedTime]
)VALUES(
    '$ID',
    '$Id_Dataset',
    '$DatasetName',
    '$Id_Workspace', 
    '$MCODETRUN',
    '$LoadToReport',
    '$TableName_cod',
    '$Table_Source',
    '$SQL_Server',
	'$SQL_Database',
	'$TableName_SQL',
	'$Columns_SQL',
    '$ExtractionDate',
    '$TableName',
	'$ModifiedTime',
    '$StructureModifiedTime'

)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_CodeM
$command.ExecuteNonQuery()
}
else {
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_insert_CodeM -AccessToken $AccessToken
}
}

# =================================================================================================================================================
# INSERT INTO -------- FAC Table ReportViews
# =================================================================================================================================================
foreach($obj in $FAC_ReportViewsJson) {
    $Id_Report                = $obj.Id_Report                    -replace "'", "''"
    $DatasetName              = $obj.DatasetName                  -replace "'", "''"
    $Date                     = $obj.Date                         -replace "'", "''"
    $DistributionMethod       = $obj.DistributionMethod           -replace "'", "''"
    $User                     = $obj.User                         -replace "'", "''"
    $Id_Workspace             = $obj.Id_Workspace                 -replace "'", "''"
    $WorkspaceName            = $obj.WorkspaceName                -replace "'", "''"
    $Report_Rank              = $obj.Report_Rank                  -replace "'", "''"

    $query_insert_tableReportViews = "INSERT INTO $tableReportViews
(
	[Id_Report],
	[DatasetName],
	[Date],
	[DistributionMethod],
	[User],
	[Id_Workspace],
	[WorkspaceName],
	[Report_Rank]
)VALUES(
	'$Id_Report',
	'$DatasetName',
	'$Date',
	'$DistributionMethod',
	'$User',
	'$Id_Workspace',
	'$WorkspaceName',
	'$Report_Rank'
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_tableReportViews
$command.ExecuteNonQuery()
}
else {
    
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_insert_tableReportViews -AccessToken $AccessToken
}
}


# =================================================================================================================================================
# INSERT INTO -------- FAC Table Database Measure
# =================================================================================================================================================
$FAC_Database_MeasureJson.Length
foreach($obj in $FAC_Database_MeasureJson) {
    $Id_Dataset               = $obj.Id_Dataset                    -replace "'", "''"
    $TableID                  = $obj.TableID                       -replace "'", "''"
    $MeasureName              = $obj.MeasureName                   -replace "'", "''"
    $MeasureType              = $obj.MeasureType                   -replace "'", "''"
    $Expression               = $obj.Expression                    -replace "'", "''"
    $LENGHT                   = if ($Expression.Length -le 7990){$Expression.Length}else {7990}
    $ExpressionTRUN           = $Expression.Substring(0, $LENGHT)    
    $ModifiedTime             = $obj.ModifiedTime                  -replace "'", "''"
    $StructureModifiedTime    = $obj.StructureModifiedTime         -replace "'", "''"
    $ErrorMessage             = $obj.ErrorMessage                  -replace "'", "''"
    $ExtractionDate           = $obj.ExtractionDate

    $query_insert_DatabaseMeasure = "INSERT INTO $tableDatabaseMeasure
(
    [Id_Dataset],
    [TableID],
    [MeasureName],
    [MeasureType],
    [Expression],        
    [ModifiedTime],
    [StructureModifiedTime],
    [ErrorMessage],
    [ExtractionDate]
)VALUES(
    '$Id_Dataset',
    '$TableID',
    '$MeasureName',
    '$MeasureType',
    '$ExpressionTRUN',        
    '$ModifiedTime',
    '$StructureModifiedTime',
    '$ErrorMessage',
    '$ExtractionDate'  
)"
if ($config.Authentication -eq "Azure Active Directory" -or $config.Authentication -eq "SQL Server Authentication")
{
$command = $connection.CreateCommand()
$command.CommandText = $query_insert_DatabaseMeasure
$command.ExecuteNonQuery()
}
else{
$AccessToken = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
Invoke-Sqlcmd -ServerInstance $Server -Database $Database  -Username $UserName -Password $Password -Query $query_insert_DatabaseMeasure -AccessToken $AccessToken
}
}



# =================================================================================================================================================
# END ----------- script
# =================================================================================================================================================
#$connection.Close()


# $DIM_Datasets
# $DIM_Dataflows
# $DIM_Workspaces
# $FAC_Dataflow_Refresh
# $DIM_Dataset_Refresh
# $FAC_Dataset_Refresh
# $DIM_Dataset_Refresh_DQ

