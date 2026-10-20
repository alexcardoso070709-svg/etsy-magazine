# Rebuild the editable magazine

The generator creates the 56-slide PPTX and the JSON page manifest. Use the supplied Fonts folder. It does not read or embed the user's reference images.

```sh
python -m pip install -r requirements.txt
python build_magazine.py
```

The default output is the parent folder of Source. **Rebuilding overwrites the generated PPTX**; work on a copy if you have edited it manually.

Optional output:

```sh
python build_magazine.py --output /path/to/another/folder
```

Copy Fonts into the new output folder first, or install the fonts in the system paths used by the generator.

PDF export requires an office application. For example:

```sh
soffice --headless --convert-to pdf --outdir ../ ../The-Days-Between-A4-Editable.pptx
```

The script uses native PowerPoint shapes, text boxes and picture placeholders. It does not flatten slides into images. PPTX generation and LibreOffice PDF export were tested in the creation environment. Microsoft PowerPoint and Canva interactive compatibility were not tested.
