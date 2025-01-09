$DiagramParameters = [xml]@'
<Diagram Name = "Core Infrastructure Services Architecture">
    <Stensils>
	    <Stensil Name="Basic" File="BASIC_M.vss" AlternativePath="false"/>
    </Stensils>
    <MasterItems>
        <MasterItem Stensil="Basic" Item="Rectangle"/>
    </MasterItems>
    <Items>
       
        <Item Type="General" Master="Rectangle" X="5.5156" Y="6.4254" Width="10.5312" Height="3.9133" FillForegnd="RGB(255,255,255)" LinePattern="0" Text="Core Infrastructure Services Architecture Diagram" VerticalAlign="0" ParaHorzAlign="0" CharSize="30 pt" CharColor="RGB(0,112,192)"/>
        <Item Type="Line" BeginX="0.2801" BeginY="7.8147" EndX="10.7812" EndY="7.8147" LineWeight="1 pt" LineColor="RGB(0,112,192)"/>
        <Item Type="General" Master="Rectangle" X="5.5156" Y="6.0852" Width="10.5312" Height="3.2328" FillForegnd="RGB(0,32,96)" LinePattern="0"/>
        <Item Type="Text" X="5.5966" Y="7.4495" Width="3.3984" Height="0.2677" Text="Core Infrastructure Services" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="2.0332" Y="6.848" Width="3.3984" Height="0.5575" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Network Time Protocol Service (NTP)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="1.1199" Y="5.8531" Width="1.5752" Height="1.2161" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Active Directory Domain Services (AD DS)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.8696" Y="5.8531" Width="1.7221" Height="1.2161" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Primary DNS" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.0332" Y="4.8668" Width="3.3984" Height="0.5575" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Secondary DNS" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.5174" Y="6.8499" Width="3.3984" Height="0.5575" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="E-Mail and Workflow" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.5174" Y="5.8531" Width="3.3984" Height="1.2161" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Active Directory Certificate Services&#xA;(AD CS)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.5174" Y="4.8668" Width="3.3984" Height="0.5575" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="IP Address Space Management (IPAM)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="8.9899" Y="6.8514" Width="3.3984" Height="0.5575" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Windows Server Update Services (WSUS)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="8.0884" Y="5.8604" Width="1.5752" Height="1.2161" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Windows Admin Center" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.8382" Y="5.8604" Width="1.7221" Height="1.2161" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Monitoring&#xA;and&#xA;Audit" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="8.9899" Y="4.8668" Width="3.3984" Height="0.5575" FillForegnd="RGB(0,176,240)" LinePattern="1" Text="Privilege and Access Management" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
    </Items>
</Diagram>
'@
.\Draw-VisioDiagram.ps1 -VCFasCodeHomeFolder "D:\VCFasCode" -DiagramFileName "CoreInfraServices.vsd" -DiagramParameters $DiagramParameters
