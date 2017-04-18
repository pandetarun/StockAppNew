Imports System.ServiceModel

' NOTE: You can use the "Rename" command on the context menu to change the interface name "IGenerateTechnicalIndicator" in both code and config file together.
<ServiceContract()>
Public Interface IGenerateTechnicalIndicator

    <OperationContract()>
    Sub GenerateIndication()

End Interface
