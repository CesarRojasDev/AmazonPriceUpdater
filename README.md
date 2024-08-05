# Automatización de Precios

Este proyecto automatiza la obtención de precios de productos desde Amazon. A partir de una columna que contiene enlaces a productos, el script utiliza Selenium para extraer los precios y guarda los resultados en un archivo Excel en el mismo directorio.

# Funcionamiento

El script se ejecuta en la terminal y solicita al usuario que ingrese los siguientes datos:

1. Nombre del archivo Excel (sin la extensión)
2. Columna del precio
3. Columna del enlace
4. Columna del stock

El script utiliza Selenium para navegar a la página de Amazon, extraer el precio y el stock y guarda los resultados en la misma hoja de cálculo.
En caso el producto no tenga stock disponible, el precio y el stock se asignan a "0" y el precio a colocar en Real Plaza sera de "10000".
## Requisitos

- Python 3.x
- [openpyxl](https://openpyxl.readthedocs.io/)
- [selenium](https://selenium-python.readthedocs.io/)
- [colorama](https://pypi.org/project/colorama/)
- WebDriver para el navegador que vayas a usar (Chrome en este caso)

Puedes instalar las librerías necesarias utilizando `pip`:

```sh
pip install openpyxl selenium colorama

```

## Ejecución

```sh
python amazon_price_updater.py
```