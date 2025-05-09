from flet import *
from datetime import datetime, timedelta
import openpyxl as xlsx
import os


class Register:
    def __init__(self, vehicle: str, day: str, datetime: datetime, action: str, quantity: int):
        self.vehicle = vehicle
        self.day = day
        self.datetime = datetime
        self.action = action
        self.quantity = quantity

    # to string
    def __str__(self):
        return f"{self.vehicle},{self.day},{self.datetime},{self.action},{self.quantity}"


def main(page: Page) -> None:
    page.client_storage.remove("selected_file")

    def pick_files_load(e: FilePickerResultEvent) -> None:
        if (e.files != None and len(e.files) > 0):
            selected_file.value = e.files[0].path
            page.client_storage.set("selected_file", e.files[0].path)
            selected_file.update()
            button_generate.visible = True
            page.update()

    def get_content_file(file_path: str) -> str:
        try:
            # Intentar abrir con UTF-8 primero
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except UnicodeDecodeError:
            try:
                # Intentar con una codificación alternativa
                with open(file_path, "r", encoding="ISO-8859-1") as f:
                    return f.read().strip()
            except UnicodeDecodeError:
                # Si también falla, probar con Windows-1252
                with open(file_path, "r", encoding="windows-1252") as f:
                    return f.read().strip()

    def verfiy_format(content_file: str) -> bool:
        lines = content_file.split("\n")
        if len(lines) == 0:
            return False
        for line in lines:
            if line.count(",") != 4:
                return False
        return True

    def convert_format(content_file: str) -> list[Register]:
        lines = content_file.split("\n")
        registers = []
        for line in lines:
            parts = line.split(",")
            date = parts[2].strip()
            # Preguntar si tiene este formato: 5/3/2024 7:00:10 p. m.
            if ("a. m." in date or "p. m." in date):
                parts[2] = date.replace("a. m.", "AM").replace("p. m.", "PM")
                # Convertir al formato 24 horas
                parts[2] = datetime.strptime(parts[2], "%d/%m/%Y %I:%M:%S %p")
                # parts[2] = parts[2].strftime("%d/%m/%Y %H:%M:%S")
            else:
                parts[2] = datetime.strptime(parts[2], "%d/%m/%Y %H:%M:%S")
                # parts[2] = parts[2].strftime("%d/%m/%Y %H:%M:%S")
            registers.append(
                Register(parts[0].strip(), parts[1].strip(), parts[2], parts[3].strip(), int(parts[4])))
        return registers

    def filter_list(list_register: list[Register]) -> list[Register]:
        i = 0
        while i < len(list_register):
            if list_register[i].quantity >= 0 and list_register[i].action == "-":
                # Buscar el anterior registro que tenga acción "+"
                j = i - 1
                while j >= 0:
                    if list_register[j].action == "+" and list_register[j].quantity > 0 and list_register[j].vehicle == list_register[i].vehicle:
                        list_register.pop(j)
                        i -= 1
                        list_register.pop(i)
                        i -= 1
                        break
                    j -= 1
            i += 1
        return list_register

    def get_types_vehicles(list_register: list[Register]) -> list[str]:
        types_vehicles = []
        for register in list_register:
            if register.vehicle not in types_vehicles:
                types_vehicles.append(register.vehicle)
        return types_vehicles

    def save_and_open_excel(wb: xlsx.Workbook, e: FilePickerResultEvent) -> None:
        path = e.path
        if not path.endswith(".xlsx"):
            path += ".xlsx"
        wb.save(path)
        # Abrir el archivo excel
        os.startfile(path)

    def pick_files_save(e: FilePickerResultEvent) -> None:
        if e.path == None or e.path == "":
            return
        selected_file_path = page.client_storage.get("selected_file")
        if selected_file_path == None:
            return

        content_file = get_content_file(selected_file_path)
        verify = verfiy_format(content_file)
        if not verify:
            return

        list_register = convert_format(content_file)
        list_register.reverse()
        filtered_list = filter_list(list_register)
        types_vehicles = get_types_vehicles(filtered_list)
        # Generar excel con las frecuencias

        wb = xlsx.Workbook()
        ws = wb.active
        ws.title = "Frecuencias"
        ws.append(["", ""] + types_vehicles)
        # Obtener el primer datetime y el último
        first_datetime = filtered_list[0].datetime
        last_datetime = filtered_list[-1].datetime
        # Hacer llegar al minuto anterior múltiplo de 5 restando con timedelta
        first_datetime = first_datetime - \
            timedelta(minutes=first_datetime.minute %
                      5, seconds=first_datetime.second)
        # Hacer llegar al minuto siguiente múltiplo de 5
        last_datetime = last_datetime + \
            timedelta(minutes=5 - last_datetime.minute %
                      5, seconds=-last_datetime.second)
        start_datetime = first_datetime
        end_datetime = last_datetime
        while start_datetime < end_datetime:
            row = []
            row.append(start_datetime.strftime("%H:%M:%S"))
            aux = start_datetime + timedelta(minutes=5)
            row.append(aux.strftime("%H:%M:%S"))
            for vehicle in types_vehicles:
                # Filtrar los registros que estén en el rango de tiempo
                registers = list(filter(
                    lambda register: register.datetime >= start_datetime and register.datetime < aux and register.vehicle == vehicle and register.action == "+" and register.quantity > 0, filtered_list))
                quantity = len(registers)

                row.append(quantity)
            start_datetime = aux
            ws.append(row)

        save_and_open_excel(wb, e)

    page.appbar = AppBar(
        title=Text("Generador de Frecuencias"),
        center_title=True
    )

    load_dialog = FilePicker(on_result=pick_files_load,)
    save_dialog = FilePicker(
        on_result=pick_files_save,

    )
    selected_file = Text()

    page.overlay.append(load_dialog)
    page.overlay.append(save_dialog)

    button_generate = ElevatedButton(
        "Generar Frecuencias",
        icon=Icons.BAR_CHART,
        visible=False,
        on_click=lambda _: save_dialog.save_file(
            allowed_extensions=["xlsx"]
        ),
    )

    page.add(
        Container(
            alignment=Alignment(0, 0),
            content=Column(
                alignment=MainAxisAlignment.CENTER,
                controls=[
                    ElevatedButton(
                        "Cargar archivo de texto",
                        icon=Icons.UPLOAD_FILE,
                        on_click=lambda _: load_dialog.pick_files(
                            allowed_extensions=["txt"]
                        ),
                    ),
                    selected_file,
                    button_generate,
                ],
            )
        )
    )


if __name__ == '__main__':
    app(target=main)
