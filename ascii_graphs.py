import math

# Simple ascii bar chart

def bar_chart(title, labels, values, width=50):
    max_val = max(values)
    print(title)
    for label, val in zip(labels, values):
        bar_len = int(val / max_val * width)
        bar = '#' * bar_len
        print(f"{label:<15} | {bar} {val}")
    print()

if __name__ == "__main__":
    products = ["FURADAN", "Juxtapid", "Alcohol"]
    # valores en miles para mantener la escala de la grafica
    quantities = [0.3, 0.553, 2100]
    bar_chart("Unidades por producto (miles)", products, quantities)

    modes = ["Maritimo", "Terrestre", "Aereo"]
    costs = [60, 30, 10]
    bar_chart("Distribucion porcentual de transporte", modes, costs)

    fases = ["Suministro", "Produccion", "Distribucion", "Cliente"]
    tiempos = [5, 8, 3, 1]
    bar_chart("Tiempos por fase (dias)", fases, tiempos)
