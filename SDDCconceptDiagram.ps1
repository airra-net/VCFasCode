# Step 1.
# Set Global Variables
$VCFasCodeHomeFolder = "D:\VCFasCode" # Folder for VCF as Code scripts and others artefacts.

# Step 2.
# Load Script Functions
Set-Location $VCFasCodeHomeFolder
. .\PSVisio.ps1 # Warning! Running scripts must be enabled on your system.

# Step 3.
# Create Visio Application
# Create Document from Blank Template
# Set Active Page
New-VisioApplication
New-VisioDocument
Set-VisioPage

# Step 4.
# Add Basic Visio Stensils
# Set Masters Item Rectangle
Add-VisioStensil -Name "Basic" -File "BASIC_M.vss"
Set-VisioStensilMasterItem -Stensil "Basic" -Item "Rectangle"

# Step 5.
# Set VMware Items Visio Stensils File Path
# Add VMware Items Visio Stensils
# Set Masters item Public Cloud, vRealize Automation, vRealize Orchestrator, VM Server, Resource Pool, vCenter Server,
# Set Masters item Rack Server, Datastore, Physical NIC
# Set Masters item Calendar
# Set Masters item vRealize Operations, vRealize log Insight, VMware Cloud Solution
# Set Masters item vCloud Availability, Site Recovery, Data Protection, Replication
# Set Masters item Secure State, Identity, Book, License
$StensilFilePath = $VCFasCodeHomeFolder + "\" + "vmw_Icons.vssx"
Add-VisioStensil -Name "VMware" -File $StensilFilePath 
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Public Cloud"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "vRealize Automation"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "vRealize Orchestrator"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "VM Server"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Resource Pool"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "vCenter Server"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Rack Server"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Datastore"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Physical NIC"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Calendar"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "vRealize Operations"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "vRealize log Insight"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "VMware Cloud Solution"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "vCloud Availability"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Site Recovery"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Data Protection"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "VR"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Secure State"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Identity"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "Book"
Set-VisioStensilMasterItem -Stensil "VMware" -Item "License"
$StensilFilePath = $VCFasCodeHomeFolder + "\" + "VMware_vCenter_Orchestrator_Shapes.vssx"
Add-VisioStensil -Name "VMwareVCO" -File $StensilFilePath 
Set-VisioStensilMasterItem -Stensil "VMwareVCO" -Item "OK"

# Step 6.
# Draw Main Rectangle, Set Size, Set Colour
# Set Header Text, Size, Color, Align
# Draw Line, Set Weight, Color
Draw-VisioItem -Master "Rectangle" -X 7.3125 -Y 7.2733 -Width 14.125 -Height 7.0467 -FillForegnd "RGB(255,255,255)" `
 -LinePattern 0 -Text "Software Defined Data Center Conceptual Diagram" -VerticalAlign 0 -ParaHorzAlign 0 `
 -CharSize "30 pt" -CharColor "RGB(0,112,192)"
Draw-VisioLine -BeginX 0.25 -BeginY 10.2294 -EndX 14.25 -EndY 10.2294 -LineWeight "1 pt" -LineColor "RGB(0,112,192)"

# Step 7. 
# Draw Cloud Automation Rectangle, Set Size, Set Colour
# Draw Text, Set Size, Color
# Draw Public Cloud, vRealize Automation, vRealize Orchestrator Icons Background Rectangle, Set Size, Set Colour
# Draw Icon Public Cloud, vRealize Automation, vRealize Orchestrator
# Draw Service Catalog Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Self-Service Portal Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Orchestration Rectangle, Set Size, Set Colour, Set Line Weight
Draw-VisioItem -Master "Rectangle" -X 2.8675 -Y 9.0947 -Width 5.2344 -Height 1.9801 -FillForegnd "RGB(152,203,225)" -LinePattern 0 
Draw-VisioText -X 1.1043 -Y 9.7391 -Width 1.7085 -Height 0.6912 -Text "Cloud Automation" -CharSize "18 pt" -CharStyle 17 -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 1.125 -Y 8.8089 -Width 1.5 -Height 1.1804 -FillForegnd "RGB(255,255,255)" -LinePattern 0 
Draw-VisioItem -Master "Public Cloud" -X 1.1043 -Y 9.0473 -Width 0.5612 -Height 0.5612
Draw-VisioItem -Master "vRealize Automation" -X 0.7436 -Y 8.5256 -Width 0.4531 -Height 0.4531
Draw-VisioItem -Master "vRealize Orchestrator" -X 1.3621 -Y 8.5256 -Width 0.4531 -Height 0.4531
Draw-VisioItem -Master "Rectangle" -X 3.6749 -Y 9.7286 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(152,203,225)" `
 -Text "Service Catalog" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt"
Draw-VisioItem -Master "Rectangle" -X 3.6749 -Y 9.1083 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(152,203,225)" `
 -Text "Self-Service Portal" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt"
Draw-VisioItem -Master "Rectangle" -X 3.6749 -Y 8.4769 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(152,203,225)" `
 -Text "Orchestration" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt"

# Step 8.  
# Draw Virtual Infrastructure Rectangle, Set Size, Set Colour
# Draw Text, Set Size, Color
# Draw VM Server, Resource Pool, vCenter Server Icons Background Rectangle, Set Size, Set Colour
# Draw Icon VM Server, Resource Pool, vCenter Server
# Draw Hypervisor Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Pools of Resources Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Virtualization Control Rectangle, Set Size, Set Colour, Set Line Weight
Draw-VisioItem -Master "Rectangle" -X 2.8675 -Y 6.9564 -Width 5.2344 -Height 1.9801 -FillForegnd "RGB(58,158,207)" -LinePattern 0 
Draw-VisioText -X 1.1043 -Y 7.6008 -Width 1.7085 -Height 0.6912 -Text "Virtual Infrastructure" -CharSize "18 pt" -CharStyle 17 -CharColor "RGB(255,255,255)" -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 1.125 -Y 6.6648 -Width 1.5 -Height 1.1804 -FillForegnd "RGB(255,255,255)" -LinePattern 0 
Draw-VisioItem -Master "VM Server" -X 1.1043 -Y 6.9285 -Width 0.5612 -Height 0.5612
Draw-VisioItem -Master "Resource Pool" -X 0.7436 -Y 6.3565 -Width 0.4531 -Height 0.4531
Draw-VisioItem -Master "vCenter Server" -X 1.3621 -Y 6.3565 -Width 0.4531 -Height 0.4531
Draw-VisioItem -Master "Rectangle" -X 3.6748 -Y 7.5903 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(58,158,207)" `
 -Text "Hypervisor" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt" -CharColor "RGB(255,255,255)" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 3.6748 -Y 6.97 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(58,158,207)" `
 -Text "Pools of Resources" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt" -CharColor "RGB(255,255,255)" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 3.6748 -Y 6.3386 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(58,158,207)" `
 -Text "Virtualization Control" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt" -CharColor "RGB(255,255,255)" -LineColor "RGB(255,255,255)"

# Step 9.
# Draw Physical Infrastructure Rectangle, Set Size, Set Colour
# Draw Text, Set Size, Color
# Draw Rack Server, Datastore, Physical NIC Icons Background Rectangle, Set Size, Set Colour
# Draw Icon Rack Server, Datastore, Physical NIC
# Draw Compute Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Storage Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Network Rectangle, Set Size, Set Colour, Set Line Weight
Draw-VisioItem -Master "Rectangle" -X 2.8625 -Y 4.8378 -Width 5.2344 -Height 1.9801 -FillForegnd "RGB(0,105,143)" -LinePattern 0 
Draw-VisioText -X 1.1043 -Y 5.4822 -Width 1.7085 -Height 0.6912 -Text "Physical Infrastructure" -CharSize "18 pt" -CharStyle 17 -CharColor "RGB(255,255,255)" -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 1.125 -Y 4.5462 -Width 1.5 -Height 1.1804 -FillForegnd "RGB(255,255,255)" -LinePattern 0
Draw-VisioItem -Master "Rack Server" -X 0.8125 -Y 4.8583 -Width 0.5612 -Height 0.1837
Draw-VisioItem -Master "Rack Server" -X 1.4694 -Y 4.8583 -Width 0.5612 -Height 0.1837
Draw-VisioItem -Master "Datastore" -X 0.789 -Y 4.375 -Width 0.4531 -Height 0.4531
Draw-VisioItem -Master "Physical NIC" -X 1.4375 -Y 4.375 -Width 0.4531 -Height 0.4531
Draw-VisioItem -Master "Rectangle" -X 3.6748 -Y 5.4717 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(0,105,143)" `
 -Text "Compute" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt" -CharColor "RGB(255,255,255)" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 3.6748 -Y 4.8514 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(0,105,143)" `
 -Text "Storage" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt" -CharColor "RGB(255,255,255)" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 3.6748 -Y 4.22 -Width 3.3828 -Height 0.5174 -FillForegnd "RGB(0,105,143)" `
 -Text "Network" -VerticalAlign 1 -ParaHorzAlign 1 `
 -CharSize "18 pt" -LinePattern 1 -LineWeight "1 pt" -CharColor "RGB(255,255,255)" -LineColor "RGB(255,255,255)"

# Step 10.
# Draw Cloud Operations Rectangle, Set Size, Set Colour
# Draw Text, Set Size, Color, Align
# Icon Calendar, OK Background Rectangle, Set Size, Set Colour
# Draw Icon Calendar, OK
# Draw Monitoring Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Monitoring Background Rectangle, Set Size, Set Colour
# Draw Icon Monitoring
# Draw Text, Set Size, Color
# Draw Logging Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Site Recovery Background Rectangle, Set Size, Set Colour
# Draw Icon Logging
# Draw Text, Set Size, Color
# Draw Life Cycle management Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Public Cloud Background Rectangle, Set Size, Set Colour
# Draw Icon Public Cloud
# Draw Masking Rectangle, Set Size, Set Colour, Set Line Weight
# Draw Icon VMware Cloud Solution
# Draw Text, Set Size, Color
Draw-VisioItem -Master "Rectangle" -X 6.9729 -Y 6.9684 -Width 2.7531 -Height 6.2451 -FillForegnd "RGB(226,232,241)" -LinePattern 0 
Draw-VisioText -X 6.9859 -Y 9.7286 -Width 1.9719 -Height 0.2878 -Text "Cloud Operations" -CharSize "18 pt" -CharStyle 17 -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 6.9674 -Y 8.6529 -Width 2.5312 -Height 1.1808 -FillForegnd "RGB(255,255,255)" -LinePattern 0 
Draw-VisioItem -Master "Calendar" -X 6.9688 -Y 8.7695 -Width 0.9375 -Height 0.8359 
Draw-VisioItem -Master "OK" -X 7.4375 -Y 8.375 -Width 0.4363 -Height 0.4363
Draw-VisioItem -Master "Rectangle" -X 6.97 -Y 7.3064 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(226,232,241)" `
  -LinePattern 1 -LineWeight "1 pt"
Draw-VisioItem -Master "Rectangle" -X 6.1352 -Y 7.3064 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "vRealize Operations" -X 6.1333 -Y 7.3064 -Width 0.4708 -Height 0.4708
Draw-VisioText -X 7.2812 -Y 7.3064 -Width 1.3725 -Height 0.2878 -Text "Monitoring" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 6.97 -Y 5.8955 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(226,232,241)" `
  -LinePattern 1 -LineWeight "1 pt"
Draw-VisioItem -Master "Rectangle" -X 6.1352 -Y 5.8955 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "vRealize log Insight" -X 6.1333 -Y 5.8788 -Width 0.4708 -Height 0.5102
Draw-VisioText -X 7.2812 -Y 5.8955 -Width 1.3725 -Height 0.2878 -Text "Logging" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 6.97 -Y 4.4952 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(226,232,241)" `
  -LinePattern 1 -LineWeight "1 pt"
Draw-VisioItem -Master "Rectangle" -X 6.1352 -Y 4.524 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "Public Cloud" -X 6.1347 -Y 4.5453 -Width 0.5431 -Height 0.5102
Draw-VisioItem -Master "Rectangle" -X 6.1438 -Y 4.3984 -Width 0.3682 -Height 0.0781 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "VMware Cloud Solution" -X 6.1408 -Y 4.4053 -Width 0.3214 -Height 0.3105
Draw-VisioText -X 7.2969 -Y 4.5227 -Width 1.7813 -Height 0.5102 -Text "Life Cycle Management" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%"

# Step 11.
# Draw Business Continuity Rectangle, Set Size, Set Colour
# Draw Text, Set Size, Color
# Icon vCloud Availability Background Rectangle, Set Size, Set Colour
# Draw Icon vCloud Availability
# Draw Fault Tolerance & Disaster Recovery Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Site Recovery Background Rectangle, Set Size, Set Colour
# Draw Icon Site Recovery
# Draw Text, Set Size, Color
# Draw Backup & Restore Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Data Protection Background Rectangle, Set Size, Set Colour
# Draw Icon Data Protection
# Draw Text, Set Size, Color
# Draw Replication Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Replication Background Rectangle, Set Size, Set Colour
# Draw Icon Replication
# Draw Text, Set Size, Color
Draw-VisioItem -Master "Rectangle" -X 9.9231 -Y 6.9741 -Width 2.7531 -Height 6.2451 -FillForegnd "RGB(29,62,125)" -LinePattern 0 
Draw-VisioText -X 9.8831 -Y 9.7286 -Width 1.2663 -Height 0.5174 -Text "Business Continuity" -CharSize "18 pt" -CharStyle 17 -CharColor "RGB(255,255,255)" -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 9.9208 -Y 8.6529 -Width 2.5312 -Height 1.1808 -FillForegnd "RGB(255,255,255)" -LinePattern 0 
Draw-VisioItem -Master "vCloud Availability" -X 9.9677 -Y 8.7447 -Width 0.956 -Height 0.8166
Draw-VisioItem -Master "Rectangle" -X 9.9232 -Y 7.3064 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(29,62,125)" `
  -LinePattern 1 -LineWeight "1 pt" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 9.1094 -Y 7.3064 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "Site Recovery" -X 9.1104 -Y 7.3007 -Width 0.4708 -Height 0.4708
Draw-VisioText -X 10.3125 -Y 7.3192 -Width 1.6875 -Height 0.9875 -Text "Fault Tolerance & Disaster Recovery" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%" -CharColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 9.9232 -Y 5.8955 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(29,62,125)" `
  -LinePattern 1 -LineWeight "1 pt" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 9.1094 -Y 5.8955 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "Data Protection" -X 9.1227 -Y 5.8822 -Width 0.4708 -Height 0.4708
Draw-VisioText -X 10.3404 -Y 5.9019 -Width 1.3725 -Height 0.5102 -Text "Backup & Restore" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%" -CharColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 9.9232 -Y 4.5227 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(29,62,125)" `
  -LinePattern 1 -LineWeight "1 pt" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 9.1094 -Y 4.5227 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "VR" -X 9.1083 -Y 4.5227 -Width 0.4708 -Height 0.4708
Draw-VisioText -X 10.3737 -Y 4.5215 -Width 1.3725 -Height 0.2878 -Text "Replication" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%" -CharColor "RGB(255,255,255)"

# Step 12.
# Draw Security and Compliance Rectangle, Set Size, Set Colour
# Draw Text, Set Size, Color
# Icon Secure State Background Rectangle, Set Size, Set Colour
# Draw Icon Secure State
# Draw Identity and Access Management Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Identity Background Rectangle, Set Size, Set Colour
# Draw Icon Identity
# Draw Text, Set Size, Color
# Draw Industry Regulations Rectangle, Set Size, Set Colour, Set Line Weight
# Icon Book Background Rectangle, Set Size, Set Colour
# Draw Icon Book
# Draw Text, Set Size, Color
# Draw Security Policies Rectangle, Set Size, Set Colour, Set Line Weight
# Icon License Background Rectangle, Set Size, Set Colour
# Draw Icon License
# Draw Text, Set Size, Color
Draw-VisioItem -Master "Rectangle" -X 12.8734 -Y 6.9741 -Width 2.7531 -Height 6.2451 -FillForegnd "RGB(100,177,69)" -LinePattern 0 
Draw-VisioText -X 12.8512 -Y 9.7286 -Width 1.4225 -Height 0.5174 -Text "Security and Compliance" -CharSize "18 pt" -CharStyle 17 -CharColor "RGB(255,255,255)" -LinePattern "0" -FillForegndTrans "100%"
Draw-VisioItem -Master "Rectangle" -X 12.8734 -Y 8.6529 -Width 2.5312 -Height 1.1808 -FillForegnd "RGB(255,255,255)" -LinePattern 0 
Draw-VisioItem -Master "Secure State" -X 12.8863 -Y 8.7447 -Width 0.956 -Height 0.8166
Draw-VisioItem -Master "Rectangle" -X 12.8607 -Y 7.3064 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(100,177,69)" `
  -LinePattern 1 -LineWeight "1 pt" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 12.0544 -Y 7.3064 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "Identity" -X 12.0505 -Y 7.3062 -Width 0.4708 -Height 0.3668
Draw-VisioText -X 13.2012 -Y 7.3192 -Width 1.6875 -Height 0.9875 -Text "Identity and Access Management" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%" -CharColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 12.8607 -Y 5.8955 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(100,177,69)" `
  -LinePattern 1 -LineWeight "1 pt" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 12.0544 -Y 5.8955 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "Book" -X 12.0521 -Y 5.8947 -Width 0.4708 -Height 0.4708
Draw-VisioText -X 13.2291 -Y 5.9019 -Width 1.3725 -Height 0.5102 -Text "Industry Regulations" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%" -CharColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 12.8607 -Y 4.5227 -Width 2.5312 -Height 1.0825 -FillForegnd "RGB(100,177,69)" `
  -LinePattern 1 -LineWeight "1 pt" -LineColor "RGB(255,255,255)"
Draw-VisioItem -Master "Rectangle" -X 12.0544 -Y 4.5227 -Width 0.5937 -Height 0.5937 -FillForegnd "RGB(255,255,255)" `
  -LinePattern 0
Draw-VisioItem -Master "License" -X 12.0625 -Y 4.5227 -Width 0.4708 -Height 0.4708
Draw-VisioText -X 13.2513 -Y 4.5215 -Width 1.3725 -Height 0.5102 -Text "Security Policies" -CharSize "18 pt" -LinePattern "0" -FillForegndTrans "100%" -CharColor "RGB(255,255,255)"

# Step 13.
# Resise Page To Fit Contents
Resize-VisioPageToFitContents

# Step 14.
# Save Document
$DiagramFileName = $VCFasCodeHomeFolder + "\" + 'SDDCConceptDiagram.vsd'
Save-VisioDocument -File $DiagramFileName

# Step 15.
# Quit Application
Close-VisioApplication






$doc = [xml]@'
<xml>
    <Section name="BackendStatus">
        <BEName BE="crust" Status="1" />
        <BEName BE="pizza" Status="1" />
        <BEName BE="pie" Status="1" />
        <BEName BE="bread" Status="1" />
        <BEName BE="Kulcha" Status="1" />
        <BEName BE="kulfi" Status="1" />
        <BEName BE="cheese" Status="1" />
    </Section>
</xml>
'@

$doc.xml.Section.BEName
$doc.xml.Section.BEName | ? { $_.Status -eq 1 }
$doc.xml.Section.BEName | ? { $_.Status -eq 1 } | % { $_.BE + " is delicious" }






$SDDCConceptualArchDiagramXml.Diagram


$SDDCConceptualArchDiagramXml.Diagram.Level | ? { $_.Name -eq "Physical Infrastructure" }


$doc.xml.Section.BEName | ? { $_.Status -eq 1 } | % { $_.BE + " is delicious" }


$doc = [xml]@'
<xml>
</xml>
'@



# Step 3.
# Set Diagram Xml Configuration
$SDDCConceptualArchDiagramXml = [xml]@'
<Diagram Name = "SDDC Conceptual Architecture">
    <Level Name = "Physical Infrastructure">
    </Level>
    <Level Name = "Virtual Infrastructure">
    </Level>
    <Level Name = "Cloud Automation">
    </Level>
    <Level Name = "Cloud Operations">
    </Level>
    <Level Name = "Business Continuity">
    </Level>
    <Level Name = "Security and Compliance">
    </Level>
</Diagram>
'@





