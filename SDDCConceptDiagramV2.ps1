$DiagramParameters = [xml]@'
<Diagram Name = "SDDC Conceptual Architecture">
    <Stensils>
	    <Stensil Name="Basic" File="BASIC_M.vss" AlternativePath="false"/>
        <Stensil Name="VMware" File="vmw_Icons.vssx" AlternativePath="true"/>
        <Stensil Name="VMwareVCO" File="VMware_vCenter_Orchestrator_Shapes.vssx" AlternativePath="true"/>
    </Stensils>
    <MasterItems>
        <MasterItem Stensil="Basic" Item="Rectangle"/>
        <MasterItem Stensil="VMware" Item="Public Cloud"/>
        <MasterItem Stensil="VMware" Item="vRealize Automation"/>
        <MasterItem Stensil="VMware" Item="vRealize Orchestrator"/>
        <MasterItem Stensil="VMware" Item="VM Server"/>
        <MasterItem Stensil="VMware" Item="Resource Pool"/>
        <MasterItem Stensil="VMware" Item="vCenter Server"/>
        <MasterItem Stensil="VMware" Item="Rack Server"/>
        <MasterItem Stensil="VMware" Item="Datastore"/>
        <MasterItem Stensil="VMware" Item="Physical NIC"/>
        <MasterItem Stensil="VMware" Item="Calendar"/>
        <MasterItem Stensil="VMware" Item="vRealize Operations"/>
        <MasterItem Stensil="VMware" Item="vRealize log Insight"/>
        <MasterItem Stensil="VMware" Item="VMware Cloud Solution"/>
        <MasterItem Stensil="VMware" Item="vCloud Availability"/>
        <MasterItem Stensil="VMware" Item="Site Recovery"/>
        <MasterItem Stensil="VMware" Item="Data Protection"/>
        <MasterItem Stensil="VMware" Item="VR"/>
        <MasterItem Stensil="VMware" Item="Secure State"/>
        <MasterItem Stensil="VMware" Item="Identity"/>
        <MasterItem Stensil="VMware" Item="Book"/>
        <MasterItem Stensil="VMware" Item="License"/>
        <MasterItem Stensil="VMwareVCO" Item="OK"/>        
    </MasterItems>
    <Items>
        <Item Type="General" Master="Rectangle" X="7.3125" Y="7.2733" Width="14.125" Height="7.0467" FillForegnd="RGB(255,255,255)" LinePattern="0" Text="Software Defined Data Center Conceptual Diagram" VerticalAlign="0" ParaHorzAlign="0" CharSize="30 pt" CharColor="RGB(0,112,192)"/>
        <Item Type="Line" BeginX="0.25" BeginY="10.229" EndX="14.25" EndY="10.2294" LineWeight="1 pt" LineColor="RGB(0,112,192)"/>
        <Item Type="General" Master="Rectangle" X="2.8675" Y="9.0947" Width="5.2344" Height="1.9801" FillForegnd="RGB(152,203,225)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="9.7391" Width="1.7085" Height="0.6912" Text="Cloud Automation" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="8.8089" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Public Cloud" X="1.1043" Y="9.047" Width="0.5612" Height="0.5612"/>
        <Item Type="General" Master="vRealize Automation" X="0.7436" Y="8.5256" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="vRealize Orchestrator" X="1.3621" Y="8.5256" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="9.7286" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Service Catalog" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="9.1083" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Self-Service Portal" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="8.4769" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Orchestration" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="2.8675" Y="6.9564" Width="5.2344" Height="1.9801" FillForegnd="RGB(58,158,207)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="7.6008" Width="1.7085" Height="0.6912" Text="Virtual Infrastructure" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="6.6648" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VM Server" X="1.1043" Y="6.9285" Width="0.5612" Height="0.5612"/>
        <Item Type="General" Master="Resource Pool" X="0.7436" Y="6.356" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="vCenter Server" X="1.3621" Y="6.3565" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="7.5903" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Hypervisor" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="6.97" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Pools of Resources" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="6.3386" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Hypervisor" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.8625" Y="4.8378" Width="5.2344" Height="1.9801" FillForegnd="RGB(0,105,143)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="5.4822" Width="1.7085" Height="0.6912" Text="Physical Infrastructure" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="4.5462" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Rack Server" X="0.8125" Y="4.8583" Width="0.5612" Height="0.1837"/>
        <Item Type="General" Master="Rack Server" X="1.4694" Y="4.8583" Width="0.5612" Height="0.1837"/>
        <Item Type="General" Master="Datastore" X="0.789" Y="4.375" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Physical NIC" X="1.4375" Y="4.375" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="5.4717" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Compute" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="4.8514" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Storage" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="4.22" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Network" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="6.9729" Y="6.9684" Width="2.7531" Height="6.2451" FillForegnd="RGB(226,232,241)" LinePattern="0"/>
        <Item Type="Text" X="6.9859" Y="9.7286" Width="1.9719" Height="0.2878" Text="Cloud Operations" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.9674" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Calendar" X="6.9688" Y="8.7695" Width="0.9375" Height="0.8359"/>
        <Item Type="General" Master="OK" X="7.4375" Y="8.375" Width="0.4363" Height="0.4363"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="7.3064" Width="0.5937" Height=" 0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vRealize Operations" X="6.1333" Y="7.3064" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="7.2812" Y="7.3064" Width="1.3725" Height="0.2878" Text="Monitoring" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vRealize log Insight" X="6.1333" Y="5.8788" Width="0.4708" Height="0.5102"/>
        <Item Type="Text" X="7.2812" Y="5.8955" Width="1.3725" Height="0.2878" Text="Logging" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="4.4952" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="4.524" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Public Cloud" X="6.1347" Y="4.5453" Width="0.5431" Height="0.5102"/>
        <Item Type="General" Master="Rectangle" X="6.1438" Y="4.3984" Width="0.3682" Height="0.0781" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VMware Cloud Solution" X="6.1408" Y="4.4053" Width="0.3214" Height="0.3105"/>
        <Item Type="Text" X="7.2969" Y="4.5227" Width="1.7813" Height="0.5102" Text="Life Cycle Management" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.9231" Y="6.9741" Width="2.7531" Height="6.2451" FillForegnd="RGB(29,62,125)" LinePattern="0"/>
        <Item Type="Text" X="9.8831" Y="9.7286" Width="1.2663" Height="0.5174" Text="Business Continuity" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="9.9208" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vCloud Availability" X="9.9677" Y="8.7447" Width="0.956" Height="0.8166"/>
        <Item Type="General" Master="Rectangle" X="9.9232" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="7.3064" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Site Recovery" X="9.1104" Y="7.3007" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3125" Y="7.3192" Width="1.6875" Height="0.9875" Text="Fault Tolerance &amp; Disaster Recovery" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.9232" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Data Protection" X="9.1227" Y="5.8822" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3404" Y="5.9019" Width="1.3725" Height="0.5102" Text="Backup &amp; Restore" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="9.9232" Y="4.5227" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="4.5227" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VR" X="9.1083" Y="4.5227" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3737" Y="4.5215" Width="1.3725" Height="0.2878" Text="Replication" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8734" Y="6.9741" Width="2.7531" Height="6.2451" FillForegnd="RGB(100,177,69)" LinePattern="0"/>
        <Item Type="Text" X="12.8512" Y="9.7286" Width="1.4225" Height="0.5174" Text="Security and Compliance" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8734" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Secure State" X="12.8863" Y="8.7447" Width="0.956" Height="0.8166"/>
        <Item Type="General" Master="Rectangle" X="12.8607" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="7.3064" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Identity" X="12.0505" Y="7.3062" Width="0.4708" Height="0.3668"/>
        <Item Type="Text" X="13.2012" Y="7.3192" Width="1.6875" Height="0.9875" Text="Identity and Access Management" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>     
        <Item Type="General" Master="Rectangle" X="12.8607" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Book" X="12.0521" Y="5.8947" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="13.2291" Y="5.9019" Width="1.3725" Height="0.5102" Text="Industry Regulations" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8607" Y="4.5227" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="4.5227" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="License" X="12.0625" Y="4.5227" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="13.2513" Y="4.5215" Width="1.3725" Height="0.5102" Text="Security Policies" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
    </Items>
</Diagram>
'@
.\Draw-VisioDiagram.ps1 -VCFasCodeHomeFolder "D:\VCFasCode" -DiagramFileName "SDDCConceptDiagramV2.vsd" -DiagramParameters $DiagramParameters
$DiagramParameters = [xml]@'
<Diagram Name = "SDDC Conceptual Architecture">
    <Stensils>
	    <Stensil Name="Basic" File="BASIC_M.vss" AlternativePath="false"/>
        <Stensil Name="VMware" File="vmw_Icons.vssx" AlternativePath="true"/>
        <Stensil Name="VMwareVCO" File="VMware_vCenter_Orchestrator_Shapes.vssx" AlternativePath="true"/>
    </Stensils>
    <MasterItems>
        <MasterItem Stensil="Basic" Item="Rectangle"/>
        <MasterItem Stensil="VMware" Item="Public Cloud"/>
        <MasterItem Stensil="VMware" Item="vRealize Automation"/>
        <MasterItem Stensil="VMware" Item="vRealize Orchestrator"/>
        <MasterItem Stensil="VMware" Item="VM Server"/>
        <MasterItem Stensil="VMware" Item="Resource Pool"/>
        <MasterItem Stensil="VMware" Item="vCenter Server"/>
        <MasterItem Stensil="VMware" Item="Rack Server"/>
        <MasterItem Stensil="VMware" Item="Datastore"/>
        <MasterItem Stensil="VMware" Item="Physical NIC"/>
        <MasterItem Stensil="VMware" Item="Calendar"/>
        <MasterItem Stensil="VMware" Item="vRealize Operations"/>
        <MasterItem Stensil="VMware" Item="vRealize log Insight"/>
        <MasterItem Stensil="VMware" Item="VMware Cloud Solution"/>
        <MasterItem Stensil="VMware" Item="vCloud Availability"/>
        <MasterItem Stensil="VMware" Item="Site Recovery"/>
        <MasterItem Stensil="VMware" Item="Data Protection"/>
        <MasterItem Stensil="VMware" Item="VR"/>
        <MasterItem Stensil="VMware" Item="Secure State"/>
        <MasterItem Stensil="VMware" Item="Identity"/>
        <MasterItem Stensil="VMware" Item="Book"/>
        <MasterItem Stensil="VMware" Item="License"/>
        <MasterItem Stensil="VMwareVCO" Item="OK"/>        
    </MasterItems>
    <Items>
        <Item Type="General" Master="Rectangle" X="7.3125" Y="7.2733" Width="14.125" Height="7.0467" FillForegnd="RGB(255,255,255)" LinePattern="0" Text="Software Defined Data Center Conceptual Diagram" VerticalAlign="0" ParaHorzAlign="0" CharSize="30 pt" CharColor="RGB(0,112,192)"/>
        <Item Type="Line" BeginX="0.25" BeginY="10.229" EndX="14.25" EndY="10.2294" LineWeight="1 pt" LineColor="RGB(0,112,192)"/>
        <Item Type="General" Master="Rectangle" X="2.8675" Y="9.0947" Width="5.2344" Height="1.9801" FillForegnd="RGB(152,203,225)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="9.7391" Width="1.7085" Height="0.6912" Text="Cloud Automation" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="8.8089" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Public Cloud" X="1.1043" Y="9.047" Width="0.5612" Height="0.5612"/>
        <Item Type="General" Master="vRealize Automation" X="0.7436" Y="8.5256" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="vRealize Orchestrator" X="1.3621" Y="8.5256" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="9.7286" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Service Catalog" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="9.1083" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Self-Service Portal" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="8.4769" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Orchestration" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="2.8675" Y="6.9564" Width="5.2344" Height="1.9801" FillForegnd="RGB(58,158,207)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="7.6008" Width="1.7085" Height="0.6912" Text="Virtual Infrastructure" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="6.6648" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VM Server" X="1.1043" Y="6.9285" Width="0.5612" Height="0.5612"/>
        <Item Type="General" Master="Resource Pool" X="0.7436" Y="6.356" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="vCenter Server" X="1.3621" Y="6.3565" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="7.5903" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Hypervisor" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="6.97" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Pools of Resources" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="6.3386" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Hypervisor" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.8625" Y="4.8378" Width="5.2344" Height="1.9801" FillForegnd="RGB(0,105,143)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="5.4822" Width="1.7085" Height="0.6912" Text="Physical Infrastructure" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="4.5462" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Rack Server" X="0.8125" Y="4.8583" Width="0.5612" Height="0.1837"/>
        <Item Type="General" Master="Rack Server" X="1.4694" Y="4.8583" Width="0.5612" Height="0.1837"/>
        <Item Type="General" Master="Datastore" X="0.789" Y="4.375" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Physical NIC" X="1.4375" Y="4.375" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="5.4717" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Compute" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="4.8514" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Storage" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="4.22" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Network" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="6.9729" Y="6.9684" Width="2.7531" Height="6.2451" FillForegnd="RGB(226,232,241)" LinePattern="0"/>
        <Item Type="Text" X="6.9859" Y="9.7286" Width="1.9719" Height="0.2878" Text="Cloud Operations" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.9674" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Calendar" X="6.9688" Y="8.7695" Width="0.9375" Height="0.8359"/>
        <Item Type="General" Master="OK" X="7.4375" Y="8.375" Width="0.4363" Height="0.4363"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="7.3064" Width="0.5937" Height=" 0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vRealize Operations" X="6.1333" Y="7.3064" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="7.2812" Y="7.3064" Width="1.3725" Height="0.2878" Text="Monitoring" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vRealize log Insight" X="6.1333" Y="5.8788" Width="0.4708" Height="0.5102"/>
        <Item Type="Text" X="7.2812" Y="5.8955" Width="1.3725" Height="0.2878" Text="Logging" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="4.4952" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="4.524" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Public Cloud" X="6.1347" Y="4.5453" Width="0.5431" Height="0.5102"/>
        <Item Type="General" Master="Rectangle" X="6.1438" Y="4.3984" Width="0.3682" Height="0.0781" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VMware Cloud Solution" X="6.1408" Y="4.4053" Width="0.3214" Height="0.3105"/>
        <Item Type="Text" X="7.2969" Y="4.5227" Width="1.7813" Height="0.5102" Text="Life Cycle Management" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.9231" Y="6.9741" Width="2.7531" Height="6.2451" FillForegnd="RGB(29,62,125)" LinePattern="0"/>
        <Item Type="Text" X="9.8831" Y="9.7286" Width="1.2663" Height="0.5174" Text="Business Continuity" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="9.9208" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vCloud Availability" X="9.9677" Y="8.7447" Width="0.956" Height="0.8166"/>
        <Item Type="General" Master="Rectangle" X="9.9232" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="7.3064" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Site Recovery" X="9.1104" Y="7.3007" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3125" Y="7.3192" Width="1.6875" Height="0.9875" Text="Fault Tolerance &amp; Disaster Recovery" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.9232" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Data Protection" X="9.1227" Y="5.8822" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3404" Y="5.9019" Width="1.3725" Height="0.5102" Text="Backup &amp; Restore" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="9.9232" Y="4.5227" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="4.5227" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VR" X="9.1083" Y="4.5227" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3737" Y="4.5215" Width="1.3725" Height="0.2878" Text="Replication" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8734" Y="6.9741" Width="2.7531" Height="6.2451" FillForegnd="RGB(100,177,69)" LinePattern="0"/>
        <Item Type="Text" X="12.8512" Y="9.7286" Width="1.4225" Height="0.5174" Text="Security and Compliance" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8734" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Secure State" X="12.8863" Y="8.7447" Width="0.956" Height="0.8166"/>
        <Item Type="General" Master="Rectangle" X="12.8607" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="7.3064" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Identity" X="12.0505" Y="7.3062" Width="0.4708" Height="0.3668"/>
        <Item Type="Text" X="13.2012" Y="7.3192" Width="1.6875" Height="0.9875" Text="Identity and Access Management" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>     
        <Item Type="General" Master="Rectangle" X="12.8607" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Book" X="12.0521" Y="5.8947" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="13.2291" Y="5.9019" Width="1.3725" Height="0.5102" Text="Industry Regulations" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8607" Y="4.5227" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="4.5227" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="License" X="12.0625" Y="4.5227" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="13.2513" Y="4.5215" Width="1.3725" Height="0.5102" Text="Security Policies" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
    </Items>
</Diagram>
'@
.\Draw-VisioDiagram.ps1 -VCFasCodeHomeFolder "D:\VCFasCode" -DiagramFileName "SDDCConceptDiagramV2.vsd" -DiagramParameters $DiagramParameters
$DiagramParameters = [xml]@'
<Diagram Name = "SDDC Conceptual Architecture">
    <Stensils>
	    <Stensil Name="Basic" File="BASIC_M.vss" AlternativePath="false"/>
        <Stensil Name="VMware" File="vmw_Icons.vssx" AlternativePath="true"/>
        <Stensil Name="VMwareVCO" File="VMware_vCenter_Orchestrator_Shapes.vssx" AlternativePath="true"/>
    </Stensils>
    <MasterItems>
        <MasterItem Stensil="Basic" Item="Rectangle"/>
        <MasterItem Stensil="VMware" Item="Public Cloud"/>
        <MasterItem Stensil="VMware" Item="vRealize Automation"/>
        <MasterItem Stensil="VMware" Item="vRealize Orchestrator"/>
        <MasterItem Stensil="VMware" Item="VM Server"/>
        <MasterItem Stensil="VMware" Item="Resource Pool"/>
        <MasterItem Stensil="VMware" Item="vCenter Server"/>
        <MasterItem Stensil="VMware" Item="Rack Server"/>
        <MasterItem Stensil="VMware" Item="Datastore"/>
        <MasterItem Stensil="VMware" Item="Physical NIC"/>
        <MasterItem Stensil="VMware" Item="Calendar"/>
        <MasterItem Stensil="VMware" Item="vRealize Operations"/>
        <MasterItem Stensil="VMware" Item="vRealize log Insight"/>
        <MasterItem Stensil="VMware" Item="VMware Cloud Solution"/>
        <MasterItem Stensil="VMware" Item="vCloud Availability"/>
        <MasterItem Stensil="VMware" Item="Site Recovery"/>
        <MasterItem Stensil="VMware" Item="Data Protection"/>
        <MasterItem Stensil="VMware" Item="VR"/>
        <MasterItem Stensil="VMware" Item="Secure State"/>
        <MasterItem Stensil="VMware" Item="Identity"/>
        <MasterItem Stensil="VMware" Item="Book"/>
        <MasterItem Stensil="VMware" Item="License"/>
        <MasterItem Stensil="VMwareVCO" Item="OK"/>        
    </MasterItems>
    <Items>
        <Item Type="General" Master="Rectangle" X="7.3125" Y="7.2733" Width="14.125" Height="7.0467" FillForegnd="RGB(255,255,255)" LinePattern="0" Text="Software Defined Data Center Conceptual Diagram" VerticalAlign="0" ParaHorzAlign="0" CharSize="30 pt" CharColor="RGB(0,112,192)"/>
        <Item Type="Line" BeginX="0.25" BeginY="10.229" EndX="14.25" EndY="10.2294" LineWeight="1 pt" LineColor="RGB(0,112,192)"/>
        <Item Type="General" Master="Rectangle" X="2.8675" Y="9.0947" Width="5.2344" Height="1.9801" FillForegnd="RGB(152,203,225)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="9.7391" Width="1.7085" Height="0.6912" Text="Cloud Automation" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="8.8089" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Public Cloud" X="1.1043" Y="9.047" Width="0.5612" Height="0.5612"/>
        <Item Type="General" Master="vRealize Automation" X="0.7436" Y="8.5256" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="vRealize Orchestrator" X="1.3621" Y="8.5256" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="9.7286" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Service Catalog" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="9.1083" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Self-Service Portal" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="3.6749" Y="8.4769" Width="3.3828" Height="0.5174" FillForegnd="RGB(152,203,225)" LinePattern="1" Text="Orchestration" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="2.8675" Y="6.9564" Width="5.2344" Height="1.9801" FillForegnd="RGB(58,158,207)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="7.6008" Width="1.7085" Height="0.6912" Text="Virtual Infrastructure" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="6.6648" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VM Server" X="1.1043" Y="6.9285" Width="0.5612" Height="0.5612"/>
        <Item Type="General" Master="Resource Pool" X="0.7436" Y="6.356" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="vCenter Server" X="1.3621" Y="6.3565" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="7.5903" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Hypervisor" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="6.97" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Pools of Resources" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="6.3386" Width="3.3828" Height="0.5174" FillForegnd="RGB(58,158,207)" LinePattern="1" Text="Hypervisor" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.8625" Y="4.8378" Width="5.2344" Height="1.9801" FillForegnd="RGB(0,105,143)" LinePattern="0"/>
        <Item Type="Text" X="1.1043" Y="5.4822" Width="1.7085" Height="0.6912" Text="Physical Infrastructure" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.125" Y="4.5462" Width="1.5" Height="1.1804" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Rack Server" X="0.8125" Y="4.8583" Width="0.5612" Height="0.1837"/>
        <Item Type="General" Master="Rack Server" X="1.4694" Y="4.8583" Width="0.5612" Height="0.1837"/>
        <Item Type="General" Master="Datastore" X="0.789" Y="4.375" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Physical NIC" X="1.4375" Y="4.375" Width="0.4531" Height="0.4531"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="5.4717" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Compute" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="4.8514" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Storage" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.6748" Y="4.22" Width="3.3828" Height="0.5174" FillForegnd="RGB(0,105,143)" LinePattern="1" Text="Network" VerticalAlign="1" ParaHorzAlign="1" CharSize="18 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="6.9729" Y="6.9684" Width="2.7531" Height="6.2451" FillForegnd="RGB(226,232,241)" LinePattern="0"/>
        <Item Type="Text" X="6.9859" Y="9.7286" Width="1.9719" Height="0.2878" Text="Cloud Operations" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.9674" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Calendar" X="6.9688" Y="8.7695" Width="0.9375" Height="0.8359"/>
        <Item Type="General" Master="OK" X="7.4375" Y="8.375" Width="0.4363" Height="0.4363"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="7.3064" Width="0.5937" Height=" 0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vRealize Operations" X="6.1333" Y="7.3064" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="7.2812" Y="7.3064" Width="1.3725" Height="0.2878" Text="Monitoring" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vRealize log Insight" X="6.1333" Y="5.8788" Width="0.4708" Height="0.5102"/>
        <Item Type="Text" X="7.2812" Y="5.8955" Width="1.3725" Height="0.2878" Text="Logging" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="6.97" Y="4.4952" Width="2.5312" Height="1.0825" FillForegnd="RGB(226,232,241)" LinePattern="1" LineWeight="1 pt"/>
        <Item Type="General" Master="Rectangle" X="6.1352" Y="4.524" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Public Cloud" X="6.1347" Y="4.5453" Width="0.5431" Height="0.5102"/>
        <Item Type="General" Master="Rectangle" X="6.1438" Y="4.3984" Width="0.3682" Height="0.0781" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VMware Cloud Solution" X="6.1408" Y="4.4053" Width="0.3214" Height="0.3105"/>
        <Item Type="Text" X="7.2969" Y="4.5227" Width="1.7813" Height="0.5102" Text="Life Cycle Management" CharSize="18 pt" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.9231" Y="6.9741" Width="2.7531" Height="6.2451" FillForegnd="RGB(29,62,125)" LinePattern="0"/>
        <Item Type="Text" X="9.8831" Y="9.7286" Width="1.2663" Height="0.5174" Text="Business Continuity" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="9.9208" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="vCloud Availability" X="9.9677" Y="8.7447" Width="0.956" Height="0.8166"/>
        <Item Type="General" Master="Rectangle" X="9.9232" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="7.3064" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Site Recovery" X="9.1104" Y="7.3007" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3125" Y="7.3192" Width="1.6875" Height="0.9875" Text="Fault Tolerance &amp; Disaster Recovery" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.9232" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Data Protection" X="9.1227" Y="5.8822" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3404" Y="5.9019" Width="1.3725" Height="0.5102" Text="Backup &amp; Restore" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="9.9232" Y="4.5227" Width="2.5312" Height="1.0825" FillForegnd="RGB(29,62,125)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.1094" Y="4.5227" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="VR" X="9.1083" Y="4.5227" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="10.3737" Y="4.5215" Width="1.3725" Height="0.2878" Text="Replication" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8734" Y="6.9741" Width="2.7531" Height="6.2451" FillForegnd="RGB(100,177,69)" LinePattern="0"/>
        <Item Type="Text" X="12.8512" Y="9.7286" Width="1.4225" Height="0.5174" Text="Security and Compliance" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8734" Y="8.6529" Width="2.5312" Height="1.1808" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Secure State" X="12.8863" Y="8.7447" Width="0.956" Height="0.8166"/>
        <Item Type="General" Master="Rectangle" X="12.8607" Y="7.3064" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="7.3064" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Identity" X="12.0505" Y="7.3062" Width="0.4708" Height="0.3668"/>
        <Item Type="Text" X="13.2012" Y="7.3192" Width="1.6875" Height="0.9875" Text="Identity and Access Management" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>     
        <Item Type="General" Master="Rectangle" X="12.8607" Y="5.8955" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="5.8955" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="Book" X="12.0521" Y="5.8947" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="13.2291" Y="5.9019" Width="1.3725" Height="0.5102" Text="Industry Regulations" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
        <Item Type="General" Master="Rectangle" X="12.8607" Y="4.5227" Width="2.5312" Height="1.0825" FillForegnd="RGB(100,177,69)" LinePattern="1" LineWeight="1 pt" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="12.0544" Y="4.5227" Width="0.5937" Height="0.5937" FillForegnd="RGB(255,255,255)" LinePattern="0"/>
        <Item Type="General" Master="License" X="12.0625" Y="4.5227" Width="0.4708" Height="0.4708"/>
        <Item Type="Text" X="13.2513" Y="4.5215" Width="1.3725" Height="0.5102" Text="Security Policies" CharSize="18 pt" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>    
    </Items>
</Diagram>
'@
.\Draw-VisioDiagram.ps1 -VCFasCodeHomeFolder "D:\VCFasCode" -DiagramFileName "SDDCConceptDiagramV2.vsd" -DiagramParameters $DiagramParameters
