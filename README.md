# Automatización de Precios

Este proyecto automatiza la obtención de precios de productos desde Amazon. A partir de una columna que contiene enlaces a productos, el script utiliza Selenium para extraer los precios y guarda los resultados en un archivo Excel en el mismo directorio.

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