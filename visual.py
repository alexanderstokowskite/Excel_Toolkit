import matplotlib.pyplot as plt

# Daten von Anscombe's Quartet
data = {
    "I": {
        "x": [4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0],
        "y": [4.26, 5.68, 6.95, 8.81, 8.04, 8.33, 9.96, 8.04, 8.84, 9.96, 7.58],
    },
    "II": {
        "x": [4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0],
        "y": [3.10, 4.74, 4.94, 5.68, 6.95, 8.81, 8.77, 9.26, 8.10, 9.13, 9.96],
    },
    "III": {
        "x": [4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0],
        "y": [5.39, 5.73, 6.08, 6.42, 6.77, 7.11, 7.46, 7.81, 8.15, 8.50, 8.84],
    },
    "IV": {
        "x": [8.0, 8.0, 8.0, 8.0, 8.0, 8.0, 8.0, 8.0, 8.0, 8.0, 19.0],
        "y": [5.25, 5.56, 5.56, 5.56, 5.56, 5.56, 5.56, 5.56, 5.56, 5.56, 12.50],
    },
}

# Erstelle die Visualisierungen
fig, axs = plt.subplots(2, 2, figsize=(10, 10))

for i, (key, value) in enumerate(data.items()):
    ax = axs[i // 2, i % 2]
    ax.scatter(value["x"], value["y"])
    ax.plot(
        value["x"],
        [sum(value["y"]) / len(value["y"])] * len(value["x"]),
        color="red",
        linestyle="--",
    )
    ax.set_title(f"Datensatz {key}")
    ax.set_xlabel("x")
    ax.set_ylabel("y")

plt.tight_layout()
plt.show()
