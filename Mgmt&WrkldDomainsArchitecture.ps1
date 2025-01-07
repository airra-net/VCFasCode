$DiagramParameters = [xml]@'
<Diagram Name = "Management and Workloads Domains Architecture">
    <Stensils>
	    <Stensil Name="Basic" File="BASIC_M.vss" AlternativePath="false"/>
    </Stensils>
    <MasterItems>
        <MasterItem Stensil="Basic" Item="Rectangle"/>
    </MasterItems>
    <Items>
        <Item Type="General" Master="Rectangle" X="5.5966" Y="4.3169" Width="10.633" Height="8.1304" FillForegnd="RGB(255,255,255)" LinePattern="0" Text="Management and Workloads Domains Architecture" VerticalAlign="0" ParaHorzAlign="0" CharSize="30 pt" CharColor="RGB(0,112,192)"/>
        <Item Type="Line" BeginX="0.2801" BeginY="7.8147" EndX="10.9193" EndY="7.8147" LineWeight="1 pt" LineColor="RGB(0,112,192)"/>
        <Item Type="Line" BeginX="0.3136" BeginY="7.6132" EndX="0.2734" EndY="7.6132" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="10.7903" BeginY="7.6132" EndX="0.3583" EndY="7.6132" LineWeight="1.5 pt" LineColor="RGB(59,158,206)" LinePattern="3"/>
        <Item Type="Line" BeginX="10.9033" BeginY="7.6132" EndX="10.8632" EndY="7.6132" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="10.9026" BeginY="7.6066" EndX="10.9026" EndY="7.5781" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="10.9026" BeginY="7.1266" EndX="10.9026" EndY="7.5366" LineWeight="1.5 pt" LineColor="RGB(59,158,206)" LinePattern="3"/>
        <Item Type="Line" BeginX="10.9026" BeginY="7.0403" EndX="10.9026" EndY="7.0122" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="10.9026" BeginY="7.0071" EndX="10.8632" EndY="7.0071" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="0.4029" BeginY="7.0071" EndX="10.835" EndY="7.0071" LineWeight="1.5 pt" LineColor="RGB(59,158,206)" LinePattern="3"/>
        <Item Type="Line" BeginX="0.3136" BeginY="7.0071" EndX="0.2734" EndY="7.0071" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="0.2734" BeginY="7.0413" EndX="0.2734" EndY="7.0132" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Line" BeginX="0.2734" BeginY="7.4935" EndX="0.2734" EndY="7.0834" LineWeight="1.5 pt" LineColor="RGB(59,158,206)" LinePattern="3"/>
        <Item Type="Line" BeginX="0.2734" BeginY="7.6063" EndX="0.2734" EndY="7.5781" LineWeight="1.5 pt" LineColor="RGB(59,158,206)"/>
        <Item Type="Text" X="5.5801" Y="7.295" Width="10.6394" Height="0.2677" Text="Another Solution Add-On" CharSize="18 pt" CharStyle="17" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="5.4583" Y="6.8192" Width="0.083" Height="0.083" FillForegnd="RGB(61,69,67)" LinePattern="0"/>
        <Item Type="General" Master="Rectangle" X="5.5966" Y="6.8192" Width="0.083" Height="0.083" FillForegnd="RGB(61,69,67)" LinePattern="0"/>
        <Item Type="General" Master="Rectangle" X="5.7349" Y="6.8192" Width="0.083" Height="0.083" FillForegnd="RGB(61,69,67)" LinePattern="0"/>
        <Item Type="General" Master="Rectangle" X="5.5966" Y="5.7842" Width="10.633" Height="1.5799" FillForegnd="RGB(9,149,212)" LinePattern="0"/>
        <Item Type="Text" X="5.6177" Y="6.3414" Width="10.539" Height="0.2677" Text="Cloud Operations and Automation Solution Add-on" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.3077" Y="5.5889" Width="1.7995" Height="0.9936" FillForegnd="RGB(9,149,212)" LinePattern="1" Text="Cross-Region Workspace ONE Access" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="3.4499" Y="5.5889" Width="1.7995" Height="0.9936" FillForegnd="RGB(9,149,212)" LinePattern="1" Text="vRealize Suite Lifecycle Manager" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.592" Y="5.5889" Width="1.7995" Height="0.9936" FillForegnd="RGB(9,149,212)" LinePattern="1" Text="vRealize Operations Manager" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="7.7342" Y="5.5889" Width="1.7995" Height="0.9936" FillForegnd="RGB(9,149,212)" LinePattern="1" Text="vRealize Log Insight" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.8763" Y="5.5889" Width="1.7995" Height="0.9936" FillForegnd="RGB(9,149,212)" LinePattern="1" Text="vRealize Automation" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.0643" Y="2.5025" Width="3.5684" Height="4.4805" FillForegnd="RGB(6,106,144)" LinePattern="0"/>
        <Item Type="Text" X="2.0546" Y="4.4439" Width="3.4126" Height="0.2677" Text="Management Domain" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="1.1501" Y="3.5605" Width="1.5752" Height="1.2161" FillForegnd="RGB(6,106,144)" LinePattern="1" Text="SDDC Manager" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.8998" Y="3.5605" Width="1.7221" Height="1.2161" FillForegnd="RGB(6,106,144)" LinePattern="1" Text="Region-Specific Workspace ONE Access" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.0616" Y="2.5845" Width="3.3984" Height="0.5575" FillForegnd="RGB(6,106,144)" LinePattern="1" Text="NSX-T" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.0616" Y="1.9378" Width="3.3984" Height="0.5575" FillForegnd="RGB(6,106,144)" LinePattern="1" Text="vCenter Server" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.0616" Y="1.2911" Width="3.3984" Height="0.5575" FillForegnd="RGB(6,106,144)" LinePattern="1" Text="vSAN" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="2.0616" Y="0.6444" Width="3.3984" Height="0.5575" FillForegnd="RGB(6,106,144)" LinePattern="1" Text="ESXi" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.6049" Y="2.5025" Width="3.024" Height="4.4805" FillForegnd="RGB(109,179,68)" LinePattern="0"/>
        <Item Type="Text" X="5.5926" Y="4.4439" Width="2.961" Height="0.2677" Text="Workload Domain" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="5.6049" Y="3.707" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="9" Text="VMware Solution for Kubernetes" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.6049" Y="2.9526" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="NSX-T (1:1 or 1:N)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.6049" Y="2.1983" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="vCenter Server" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.6049" Y="1.4439" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="Shared Storage&#xA;(vSAN, NFS, VMFS)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="5.6049" Y="0.6895" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="ESXi" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="7.3647" Y="2.5025" Width="0.083" Height="0.083" FillForegnd="RGB(61,69,67)" LinePattern="0"/>
        <Item Type="General" Master="Rectangle" X="7.503" Y="2.5025" Width="0.083" Height="0.083" FillForegnd="RGB(61,69,67)" LinePattern="0"/>
        <Item Type="General" Master="Rectangle" X="7.6413" Y="2.5025" Width="0.083" Height="0.083" FillForegnd="RGB(61,69,67)" LinePattern="0"/>
        <Item Type="General" Master="Rectangle" X="9.4011" Y="2.5025" Width="3.024" Height="4.4805" FillForegnd="RGB(109,179,68)" LinePattern="0"/>
        <Item Type="Text" X="9.4068" Y="4.4439" Width="3.0111" Height="0.2677" Text="Workload Domain" CharSize="18 pt" CharStyle="17" CharColor="RGB(255,255,255)" LinePattern="0" FillForegndTrans="100%"/>
        <Item Type="General" Master="Rectangle" X="9.4011" Y="3.707" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="9" Text="VMware Solution for Kubernetes" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.4011" Y="2.9526" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="NSX-T (1:1 or 1:N)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.4011" Y="2.1983" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="vCenter Server" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.4011" Y="1.4439" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="Shared Storage&#xA;(vSAN, NFS, VMFS)" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
        <Item Type="General" Master="Rectangle" X="9.4011" Y="0.6895" Width="2.8094" Height="0.6652" FillForegnd="RGB(109,179,68)" LinePattern="1" Text="ESXi" VerticalAlign="1" ParaHorzAlign="1" CharSize="14 pt" LineWeight="1 pt" CharColor="RGB(255,255,255)" LineColor="RGB(255,255,255)"/>
    </Items>
</Diagram>
'@
.\Draw-VisioDiagram.ps1 -VCFasCodeHomeFolder "D:\VCFasCode" -DiagramFileName "Mgmt&WrkldDomainsArchitecture.vsd" -DiagramParameters $DiagramParameters
