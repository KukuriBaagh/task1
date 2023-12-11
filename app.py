from flask import Flask, request, render_template
import openpyxl

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/calculate", methods=["POST"])
def calculate():
    try:
        number1 = float(request.form["number1"])
        number2 = float(request.form["number2"])

        result = number1 + number2

        # Writing to Excel sheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet["A1"] = "Number 1"
        sheet["B1"] = "Number 2"
        sheet["C1"] = "Result"
        sheet.append([number1, number2, result])

        workbook.save("calculation_result.xlsx")

        return f"Calculation successful. Result: {result}. Check the Excel sheet for details."

    except Exception as e:
        return f"Error: {str(e)}"


if __name__ == "__main__":
    app.run(debug=True)
