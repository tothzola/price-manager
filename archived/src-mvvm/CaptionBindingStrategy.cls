VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaptionBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A target binding strategy that creates a CaptionPropertyBinding by default, or a OneWayPropertyBinding given a target property path."
'@ModuleDescription "A target binding strategy that creates a CaptionPropertyBinding by default, or a OneWayPropertyBinding given a target property path."
'@Folder("Structure.Bindings.Strategies")
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IBindingTargetStrategy

Private Function IBindingTargetStrategy_DefaultPropertyBindingFor(ByVal Context As IContextModel, ByVal Source As IBindingPath, ByVal Target As Object, _
                                                                  ByVal Mode As BindingMode, _
                                                                  ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
                                                                  ByVal StringFormat As IStringFormatter) As IPropertyBinding

    Set IBindingTargetStrategy_DefaultPropertyBindingFor = CaptionPropertyBinding.Create(Context, Source, Target, _
                                                                                         StringFormat:=StringFormat)

End Function

Private Function IBindingTargetStrategy_PropertyBindingFor(ByVal Context As IContextModel, ByVal Source As IBindingPath, ByVal Target As IBindingPath, _
                                                           ByVal Mode As BindingMode, _
                                                           ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
                                                           ByVal StringFormat As IStringFormatter) As IPropertyBinding

    Set IBindingTargetStrategy_PropertyBindingFor = OneWayPropertyBinding.Create(Context, Source, Target, _
                                                                                 StringFormat:=StringFormat)
                                                                                 
End Function


