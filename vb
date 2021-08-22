//get selected item dropdown

<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
	<ribbon>
		<tabs>
			<tab id="taxes" label="Taxes" insertAfterMso="TabHome">
				<group idMso="GroupClipboard" />
				<group idMso="GroupFont" />
				<group id="customGroup" label="Contoso Tools">
					<button id="customButton1" label="ConBold" size="large" onAction="conBoldSub" imageMso="Bold" />
					<button id="customButton2" label="ConItalic" size="large" onAction="conItalicSub" imageMso="Italic" />
					<button id="customButton3" label="ConUnderline" size="large" onAction="conUnderlineSub" imageMso="Underline" />
				</group>
				<group id="indicadores" label="Indicadores">
					<dropDown id="dropDown" label="Indicador" onAction="DDOnAction" >
   						<item id="item1" label="Item 1" />
   						<item id="item2" label="Item 2" />
   						<item id="item3" label="Item 3" />
   						<button id="button" label="Button..." />
 					</dropDown>
				</group>
				<group idMso="GroupEnterDataAlignment" />
				<group idMso="GroupEnterDataNumber" />
				<group idMso="GroupQuickFormatting" />
			</tab>
		</tabs>
	</ribbon>
</customUI>






Sub DDOnAction(control As IRibbonControl, id As String, Index As Integer)
Select Case Index
    Case 0
        Debug.Print "hola"
    Case 1
        ActiveWorkbook.Sheets("Sheet2").Activate
    Case 2
        ActiveWorkbook.Sheets("Sheet3").Activate
End Select

End Sub
