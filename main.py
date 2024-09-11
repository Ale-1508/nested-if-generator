# flet pack --product-name "Nested If Generator" --file-description "Nested IF excel formula generator" --product-version "1.2" --file-version "0.0.1.2" --company-name "IG3.SRL" --copyright "Alessandro M" -i "icon.ico" .\SistemaExcel.py
import flet as ft
import subprocess
from typing import Union, Optional, Callable, Any

BGCOLOR = "#efeff0"

class State:
    def __init__(self):
        self.targetCell:TextfieldSmall = None
        self.defautlValue:TextfieldSmall = None
        self.conditions:list[Condition] = []
        self.addConditionToList:Optional[Callable] = None
        self.deleteConditionFromList:Optional[Callable[[str], None]] = None

    def add(self, data:list[tuple]):
        for element in data:
            if len(element)!=2: continue
            self.__setattr__(element[0], element[1])

appState = State()

class Button(ft.ElevatedButton):
    def __init__(self, 
                 text:str, 
                 data:Optional[str]=None, 
                 tooltip:Optional[str]=None,
                 width:int=196, 
                 height:int=None, 
                 bgcolor:str=ft.colors.WHITE70, 
                 text_color:str=ft.colors.BLACK, 
                 on_click=Callable[[Any], Any],
                ):
        super().__init__(
            text=text,
            data=data,
            tooltip=tooltip,
            width=width,
            height=height,
            bgcolor=bgcolor,
            color=text_color,
            style=ft.ButtonStyle(
                shape={
                    ft.ControlState.DEFAULT : ft.RoundedRectangleBorder(radius=8),
                    ft.ControlState.HOVERED : ft.RoundedRectangleBorder(radius=12),
                },
            ),
            col={"xs":6,"sm":4,"md":3},
            on_click=on_click,
        )

    def update(self, text:str):
        self.text = text
        super().update()
    
class Text(ft.Text):
    def __init__(self, text:str="",color:str=ft.colors.WHITE, size:Optional[int]=None):
        super().__init__(
            value=text,
            color=color,
            size=size,
            weight=ft.FontWeight.W_500
        )

class TextfieldSmall(ft.Container):
    def __init__(
            self, 
            value:str="",
            hint_text:str="",
            col:Optional[dict]=None,
            width:Optional[int]=None,
            height:Optional[int]=None,
            expand:Union[int,bool]=False,
            bgcolor:str=ft.colors.WHITE,
            alignment:ft.alignment=ft.alignment.center_left,
        ):
        super().__init__(
            col=col,
            height=height,
            expand=expand,
            bgcolor=bgcolor,
            border_radius=12,
            alignment=alignment,
            margin=ft.margin.only(right=8),
            padding=ft.padding.symmetric(horizontal=8),
            content=ft.TextField(
                dense=True,
                value=value,
                hint_text=hint_text,
                border=ft.InputBorder.NONE,
            ),
        )

    def reset(self):
        self.content.value = ""
        self.content.update()

    @property
    def value(self):
        return self.content.value

class HeaderContainer(ft.Container):
    textfieldStyle = { 
        "bgcolor": BGCOLOR,
        "height": 36
    }
    def __init__(self):
        targetCell = TextfieldSmall(hint_text="Cell", col="4", **self.textfieldStyle)
        defautlValue = TextfieldSmall(hint_text="Default Value", col="8", **self.textfieldStyle)
        
        appState.add([
            ("targetCell", targetCell),
            ("defautlValue", defautlValue)
        ])
        
        super().__init__(
            content=ft.ResponsiveRow(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                spacing=12,
                controls=[
                    targetCell,
                    defautlValue
                ],
            ),
        )

class MainContainer(ft.Container):
    def __init__(self):
        super().__init__(
            expand=True,
            content=ft.Column(
                spacing=16,
                controls=[
                    AddCondition(),
                    Text("Conditions List:", color="white", size=18),
                    ConditionsList(),
                ],
            ),
        )

class AddCondition(ft.Container):
    textfieldStyle = { 
        "bgcolor": BGCOLOR,
        "height": 36
    }
    def __init__(self):
        self.condition = TextfieldSmall(hint_text="Condition", col="6", **self.textfieldStyle)
        self.ifTrue = TextfieldSmall(hint_text="If True", col="6", **self.textfieldStyle)

        super().__init__(
            padding=16,
            border_radius=16,
            bgcolor="red",
            gradient=ft.LinearGradient(
                [
                    ft.colors.BLUE_800,
                    ft.colors.BLUE_600,
                ],
                begin = ft.alignment.center_left,
                end = ft.alignment.center_right,
            ),
            content=ft.Column(
                spacing=16,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    HeaderContainer(), 
                    ft.ResponsiveRow(
                        controls=[
                            self.condition,
                            self.ifTrue,
                        ],
                    ),
                    Button("Add Condition", on_click=self.add),
                ]
            )
        )

    def add(self, e:ft.TapEvent):
        condition = {
            "condition":self.condition.value,
            "ifTrue":self.ifTrue.value,
        }
        appState.addConditionToList(condition)
        for textfield in [self.condition, self.ifTrue]: textfield.reset()

class Condition(ft.Container):
    def __init__(self, condition:str, ifTrue:str):
        self.condition = condition
        self.ifTrue = ifTrue

        super().__init__(
            expand=True,
            border_radius=12,
            col={
                "xs":12,
                "sm":12, 
                "md":6, 
                "lg":4, 
            },
            gradient=ft.LinearGradient(
                [
                    ft.colors.BLUE_800,
                    ft.colors.BLUE_600,
                ],
                begin = ft.alignment.center_left,
                end = ft.alignment.center_right,
            ),
            alignment=ft.alignment.center_left,
            padding=ft.padding.symmetric(vertical=8, horizontal=16),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                controls=[
                    ft.Text(f"""IF( {appState.targetCell.value.upper()} = "{condition}"; "{ifTrue}"; ... )""", color="#F5F5F5"),
                    ft.IconButton(
                        icon=ft.icons.REMOVE_ROUNDED, 
                        icon_color="#F5F5F5", 
                        on_click=self.handleDelete
                    ),
                ]
            ),
        )

    def handleDelete(self, e:ft.TapEvent):
        appState.deleteConditionFromList(self.data)

class ConditionsList(ft.Container):
    def __init__(self):
        self.conditionID = 0
        self.conditions = appState.conditions
        appState.addConditionToList = self.addCondition
        appState.deleteConditionFromList = self.deleteCondition
        super().__init__(
            expand=True,
            alignment=ft.alignment.top_center,
            content=ft.Column(
                scroll=ft.ScrollMode.HIDDEN,
                controls=[
                    ft.ResponsiveRow(    
                        controls=self.conditions
                    ),
                ]
            )
        )

    def addCondition(self, condition:dict):
        if checkTextfieldNullOrEmpty(appState.targetCell.value, self.page, "Insert a Cell"): return
        try: int(appState.targetCell.value[-1])
        except ValueError: raiseError(self.page, "Invalid cell format")
        else:
            conditionContainer = Condition(**condition)
            conditionContainer.data = self.conditionID
            self.conditionID += 1
            self.content.controls[0].controls.append(conditionContainer)
            appState.conditions.append(conditionContainer)
            self.update()

    def deleteCondition(self, conditionID:dict):
        for index, condition in enumerate(appState.conditions):
            if not conditionID == condition.data: continue
            self.content.controls[0].controls.pop(index)
            appState.conditions.pop(index)
            self.update()
            return

def raiseError(page:ft.Page, message:str): openDialog(page, ft.AlertDialog(title=ft.Text(message, text_align=ft.TextAlign.CENTER)))

def checkTextfieldNullOrEmpty(value:str, page:ft.Page, error:str="Empty Field")->bool:
    if not value == "" or value: return False
    raiseError(page, error)
    return True

def copyToClipboard(text): subprocess.run("clip", text=True, input=text)

def openDialog(page: ft.Page, dialog:ft.AlertDialog):
    page.overlay.append(dialog)
    dialog.open = True
    page.update()

def generateFormula(page: ft.Page):
    formula = f'"{appState.defautlValue.value}"'
    for condition in appState.conditions:
        formula = f"""IF({appState.targetCell.value.upper()}="{condition.condition}";"{condition.ifTrue}";{formula})"""
    copyToClipboard("=" + formula)
    openDialog(page, ft.AlertDialog(title=ft.Text("Generated and Copied Formula!", text_align=ft.TextAlign.CENTER)))

def main(page: ft.Page):
    page.title = 'Nested IF Generator'
    page.window.width = 480
    page.window.height = 560
    page.window.center()

    page.padding = 0
    page.window.bgcolor= ft.colors.TRANSPARENT
    
    page.horizontal_alignment = ft.MainAxisAlignment.CENTER
    page.vertical_alignment = ft.CrossAxisAlignment.CENTER
    page.add(
        ft.Container(
            expand=True,
            bgcolor="#303030",
            padding = ft.padding.symmetric(vertical=16, horizontal=32),
            content=ft.Column(
                expand=True,
                alignment=ft.MainAxisAlignment.START,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    MainContainer(),
                    Button("Generate Formula", on_click=lambda e: generateFormula(page))]
            )
        )    
    )

if __name__ == '__main__':
    ft.app(target=main)