from analyzer import analizar_documento

def comparar_documentos(ruta1, ruta2):
    resultado1 = analizar_documento(ruta1)
    resultado2 = analizar_documento(ruta2)

    comparacion = {
        "version1": {
            "nombre": resultado1["nombre_archivo"],
            "puntaje": resultado1["puntaje"]
        },
        "version2": {
            "nombre": resultado2["nombre_archivo"],
            "puntaje": resultado2["puntaje"]
        },
        "diferencia_puntaje": resultado2["puntaje"] - resultado1["puntaje"],
        "dimensiones": []
    }

    for dim1, dim2 in zip(resultado1["dimensiones"], resultado2["dimensiones"]):
        items_v1 = {i["titulo"]: i for i in dim1["items"]}
        items_v2 = {i["titulo"]: i for i in dim2["items"]}

        corregidos = []
        persisten = []
        regresiones = []
        nuevos_ok = []

        todos_titulos = set(items_v1.keys()) | set(items_v2.keys())

        for titulo in todos_titulos:
            i1 = items_v1.get(titulo)
            i2 = items_v2.get(titulo)

            if i1 and i2:
                if i1["estado"] in ["error", "advertencia"] and i2["estado"] == "ok":
                    corregidos.append({"titulo": titulo, "antes": i1["estado"], "detalle": i2["detalle"]})
                elif i1["estado"] == "ok" and i2["estado"] in ["error", "advertencia"]:
                    regresiones.append({"titulo": titulo, "despues": i2["estado"], "detalle": i2["detalle"], "sugerencia": i2["sugerencia"]})
                elif i1["estado"] in ["error", "advertencia"] and i2["estado"] in ["error", "advertencia"]:
                    persisten.append({"titulo": titulo, "estado": i2["estado"], "detalle": i2["detalle"], "sugerencia": i2["sugerencia"]})
            elif not i1 and i2:
                if i2["estado"] == "ok":
                    nuevos_ok.append({"titulo": titulo, "detalle": i2["detalle"]})
                else:
                    regresiones.append({"titulo": titulo, "despues": i2["estado"], "detalle": i2["detalle"], "sugerencia": i2["sugerencia"]})

        comparacion["dimensiones"].append({
            "nombre": dim1["nombre"],
            "puntaje_v1": dim1["puntaje"],
            "puntaje_v2": dim2["puntaje"],
            "corregidos": corregidos,
            "persisten": persisten,
            "regresiones": regresiones,
            "nuevos_ok": nuevos_ok
        })

    return comparacion
