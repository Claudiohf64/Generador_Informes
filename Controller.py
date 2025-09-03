import os
import tempfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, redirect, url_for, flash

from ModeloInforme import generar_informe
from ModeloCuadro import generar_cuadro, generar_descripciones_mistral

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "devkey")  # Seguro en producción


# === Página principal ===
@app.route("/")
def index():
    return render_template("index.html")


# === Generar Informe ===
@app.route("/generar_informe", methods=["POST"])
def generar_informe_view():
    try:
        file = request.files.get("archivo_base")
        if not file or file.filename == "":
            flash("Debes subir un archivo base para el informe")
            return redirect(url_for("index"))

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_base:
            file.save(temp_base.name)

        tarea = request.form.get("tarea", "Tarea no especificada")
        puntos_texto = request.form.get("puntos", "").strip()
        puntos = [p.strip() for p in puntos_texto.splitlines() if p.strip()]

        if not puntos:
            flash("Debes ingresar al menos un punto")
            os.unlink(temp_base.name)
            return redirect(url_for("index"))

        # Archivo de salida temporal
        salida = tempfile.mktemp(suffix=".docx")
        ancla = "[[INICIO_INFORME]]"

        path_final = generar_informe(temp_base.name, salida, puntos, tarea, ancla=ancla)
        os.unlink(temp_base.name)

        return send_file(path_final, as_attachment=True)

    except Exception as e:
        flash(f"Error al generar informe: {e}")
        return redirect(url_for("index"))


# === Generar Cuadro ===
@app.route("/generar_cuadro", methods=["POST"])
def generar_cuadro_view():
    try:
        file = request.files.get("archivo_base_cuadro")
        if not file or file.filename == "":
            flash("Debes subir un archivo base para el cuadro")
            return redirect(url_for("index"))

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_base:
            file.save(temp_base.name)

        dias_semana = []
        for i in range(1, 7):
            dia = request.form.get(f"dia_{i}")
            fecha = request.form.get(f"fecha_{i}")
            laborable = request.form.get(f"laborable_{i}")
            hora_inicio = request.form.get(f"hora_inicio_{i}")
            hora_fin = request.form.get(f"hora_fin_{i}")
            tema = request.form.get(f"tema_{i}")
            tareas_texto = request.form.get(f"tareas_{i}", "")

            # Día vacío → no laborable
            if not fecha and not tema and not tareas_texto.strip():
                dias_semana.append({
                    "dia": dia,
                    "fecha": "",
                    "laborable": False,
                    "razon_no_lab": "Día sin actividades",
                    "temas": [],
                    "horas": "",
                })
                continue

            if laborable == "si":
                horas = f"{hora_inicio}–{hora_fin}" if hora_inicio and hora_fin else ""
                tareas = [t.strip() for t in tareas_texto.splitlines() if t.strip()]

                try:
                    tareas_con_desc = generar_descripciones_mistral(tema or "Sin tema", tareas) if tareas else []
                except Exception:
                    tareas_con_desc = [{"nombre": t, "descripcion": "Realicé una tarea sin descripción"} for t in tareas]

                temas = [{"tema": tema or "Sin tema", "tareas": tareas_con_desc}] if tareas_con_desc else []

                dias_semana.append({
                    "dia": dia,
                    "fecha": fecha or "",
                    "laborable": True,
                    "temas": temas,
                    "horas": horas,
                })
            else:
                dias_semana.append({
                    "dia": dia,
                    "fecha": fecha or "",
                    "laborable": False,
                    "razon_no_lab": "Día no laborable",
                    "temas": [],
                    "horas": "",
                })

        salida = tempfile.mktemp(suffix=".docx")
        ancla = "[[AQUI_TABLA]]"

        path_final = generar_cuadro(temp_base.name, salida, dias_semana, ancla=ancla)
        os.unlink(temp_base.name)

        return send_file(path_final, as_attachment=True)

    except Exception as e:
        flash(f"Error al generar cuadro: {e}")
        return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
