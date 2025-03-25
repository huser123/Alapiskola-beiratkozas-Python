import os
import email
import re
import csv
import html
import quopri
from bs4 import BeautifulSoup
from email.header import decode_header
import codecs

def decode_email_header(header):
   """Dekódolja az email fejléceket."""
   decoded_parts = []
   for part, encoding in decode_header(header):
       if isinstance(part, bytes):
           if encoding:
               try:
                   decoded_parts.append(part.decode(encoding))
               except:
                   decoded_parts.append(part.decode('utf-8', errors='replace'))
           else:
               decoded_parts.append(part.decode('utf-8', errors='replace'))
       else:
           decoded_parts.append(part)
   return " ".join(decoded_parts)

def extract_field_value(html_content, field_name):
   """Kinyeri a mező értékét a HTML tartalomból."""
   try:
       soup = BeautifulSoup(html_content, 'html.parser')
       
       # Keresni a mezőnevet a <strong> tagekben
       strong_tags = soup.find_all('strong', string=lambda s: field_name in s if s else False)
       
       if strong_tags:
           # A mező értéke a következő td.field-value tagben található
           for strong in strong_tags:
               parent_tr = strong.find_parent('tr')
               if parent_tr:
                   next_tr = parent_tr.find_next_sibling('tr')
                   if next_tr:
                       value_td = next_tr.find('td', class_='field-value')
                       if value_td:
                           # Ellenőrizni, hogy van-e benne <a> tag (email címek esetén)
                           if value_td.find('a'):
                               return value_td.find('a').text
                           # Több soros checkboxok esetén
                           if value_td.find('br'):
                               values = []
                               for content in value_td.contents:
                                   if content.name != 'br':
                                       values.append(content.strip())
                               return ", ".join([v for v in values if v])
                           return value_td.text
   except Exception as e:
       print(f"Hiba a mező kinyerésénél: {field_name} - {e}")
   return ""

def process_eml_file(eml_path):
   """Feldolgoz egy .eml fájlt és visszaadja a kinyert adatokat."""
   try:
       with open(eml_path, 'rb') as f:
           msg = email.message_from_binary_file(f)
       
       # HTML tartalom kinyerése
       html_content = ""
       if msg.is_multipart():
           for part in msg.walk():
               if part.get_content_type() == "text/html":
                   payload = part.get_payload(decode=True)
                   charset = part.get_content_charset() or 'utf-8'
                   try:
                       html_content = payload.decode(charset)
                   except UnicodeDecodeError:
                       html_content = payload.decode('utf-8', errors='replace')
                   break
       else:
           payload = msg.get_payload(decode=True)
           charset = msg.get_content_charset() or 'utf-8'
           try:
               html_content = payload.decode(charset)
           except UnicodeDecodeError:
               html_content = payload.decode('utf-8', errors='replace')
       
       # Quoted-printable dekódolás
       try:
           html_content = quopri.decodestring(html_content.encode('utf-8')).decode('utf-8')
       except Exception as e:
           print(f"Hiba a quoted-printable dekódolás során: {e}")
           # Ha hiba történik, megtartjuk az eredeti tartalmat
           pass
       
       # HTML entitások dekódolása
       html_content = html.unescape(html_content)
       
       # Adatok kinyerése
       data = {}
       
       # Alap mezők
       fields = [
           "A tanuló neve:",
           "A tanuló születési helye és dátuma:",
           "A tanuló születési száma:",
           "A tanuló állandó lakhelye:",
           "A tanuló állampolgársága:",
           "A tanuló nemzetisége:",
           "Az apa neve:",
           "Az apa e-mail címe:",
           "Az apa telefonszáma:",
           "Az apa állandó lakhelye:",
           "Az anya neve:",
           "Az anya e-mail címe:",
           "Az anya telefonszáma:",
           "Az anya állandó lakhelye:",
           "Melyik óvodába járt a gyermek:",
           "Milyen jellegű osztályt választana (több lehetőség is választható):",
           "Választható tantárgy:",
           "Napköziotthon:",
           "Iskolai étkeztetésre igényt tart:",
           "Van a gyermekének allergiája",
           "A szülők egy háztartásban élnek?",
           "Elsődleges kapcsolattartási személy vezetékneve és keresztneve",
           "Elsődleges kapcsolattartási telefonszám",
           "Elsődleges kapcsolattartási e-mail cím",
           "Az iskolalátogatásról szóló határozatot"
       ]
       
       for field in fields:
           value = extract_field_value(html_content, field)
           # Tisztítás
           if value:
               value = value.strip()
               # Mező név átalakítások
               if field == "Van a gyermekének allergiája":
                   field = "Van a gyermekének allergiája vagy más betegsége, melyről az iskolálak tudnia kell?:"
               if field == "Az iskolalátogatásról szóló határozatot":
                   field = "Az iskolalátogatásról szóló határozatot, illetve más levelezést az iskola kinek a nevére címezheti?:"
               if field == "Elsődleges kapcsolattartási személy vezetékneve és keresztneve":
                   field = "Elsődleges kapcsolattartási személy vezetékneve és keresztneve (kit kereshetünk iskolai ügyekben):"
               if field == "Elsődleges kapcsolattartási telefonszám":
                   field = "Elsődleges kapcsolattartási telefonszám (kit kereshetünk iskolai ügyekben):"
               if field == "Elsődleges kapcsolattartási e-mail cím":
                   field = "Elsődleges kapcsolattartási e-mail cím (kit kereshetünk iskolai ügyekben):"
               
           data[field] = value
       
       # Ellenőrizzük, hogy nincs-e esetleg "Bármilyen egyéb megjegyzés" mező
       megjegyzes = extract_field_value(html_content, "Bármilyen egyéb megjegyzés")
       if megjegyzes:
           data["Bármilyen egyéb megjegyzés, amiről esetleg tudnunk kellene:"] = megjegyzes
       else:
           data["Bármilyen egyéb megjegyzés, amiről esetleg tudnunk kellene:"] = ""
       
       # Debug: kiírjuk a kinyert adatokat
       print("Kinyert adatok:")
       for field, value in data.items():
           print(f"  {field}: {value}")
       
       return data
   except Exception as e:
       print(f"Hiba a fájl feldolgozása közben: {str(e)}")
       raise

def escape_latex(text):
   """TeX-biztos kódolás"""
   if not text:
       return ""
   text = str(text)
   # Speciális LaTeX karakterek escape-elése
   replacements = {
       '&': r'\&',
       '%': r'\%',
       '$': r'\$',
       '#': r'\#',
       '_': r'\_',
       '{': r'\{',
       '}': r'\}',
       '~': r'\textasciitilde{}',
       '^': r'\textasciicircum{}',
       '\\': r'\textbackslash{}',
   }
   for char, replacement in replacements.items():
       text = text.replace(char, replacement)
   return text

def extract_dataentry_template(template_path):
    """Kivonja az adatlap sablont a teljes TeX sablonfájlból."""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            tex_content = f.read()
        
        # Keresünk egy minta adatlapot a !!!VARIABLE!!! említések között
        # Feltételezzük, hogy a \begin{document} és \end{document} között van
        begin_doc = tex_content.find(r'\begin{document}')
        end_doc = tex_content.find(r'\end{document}')
        
        if begin_doc == -1 or end_doc == -1:
            raise ValueError("Nem található \\begin{document} vagy \\end{document} a sablonban")
        
        # Az első változó előfordulása
        first_var = tex_content.find('!!!VARIABLE!!!', begin_doc)
        if first_var == -1 or first_var > end_doc:
            raise ValueError("Nem található !!!VARIABLE!!! a dokumentum törzsében")
        
        # Az utolsó változó előfordulása
        last_var = tex_content.rfind('!!!VARIABLE!!!', begin_doc, end_doc)
        
        # Keresünk egy teljes adatlapot (az első változótól az utolsóig)
        # majd kibővítjük a következő \newpage-ig vagy a dokumentum végéig
        entry_start = tex_content.rfind(r'\begin{figure}', begin_doc, first_var)
        if entry_start == -1:
            entry_start = tex_content.rfind(r'\begin{center}', begin_doc, first_var)
        if entry_start == -1:
            entry_start = tex_content.rfind(r'\begin{tabular}', begin_doc, first_var)
        if entry_start == -1:
            # Ha nem találtunk jelölőt, visszalépünk 50 karaktert az első változó előtt
            entry_start = max(begin_doc + len(r'\begin{document}'), first_var - 50)
        
        entry_end = tex_content.find(r'\newpage', last_var)
        if entry_end == -1 or entry_end > end_doc:
            entry_end = end_doc
        
        # Kivonjuk a dokumentumból a preambulumot, egy adatlapot és a lezárást
        preamble = tex_content[:begin_doc + len(r'\begin{document}')]
        entry_template = tex_content[entry_start:entry_end + len(r'\newpage')]
        footer = tex_content[end_doc:]
        
        print("Sikeresen kivonat a preambulum, adatlap sablon és lábléc szakaszokat.")
        
        return preamble, entry_template, footer
    except Exception as e:
        print(f"Hiba a sablonfájl feldolgozása közben: {str(e)}")
        raise

def create_tex_file(all_data, output_path, template_path):
    """Létrehozza a TeX fájlt az adatokkal a külső sablonfájlból."""
    try:
        # Kivonjuk a sablonfájl részeit
        preamble, entry_template, footer = extract_dataentry_template(template_path)
        
        # A mezők sorrendje a sablonfájlban
        field_order = [
            "A tanuló neve:",
            "A tanuló születési helye és dátuma:",
            "A tanuló születési száma:",
            "A tanuló állandó lakhelye:", 
            "A tanuló állampolgársága:",
            "A tanuló nemzetisége:",
            "Az apa neve:",
            "Az apa e-mail címe:",
            "Az apa telefonszáma:",
            "Az apa állandó lakhelye:",
            "Az anya neve:",
            "Az anya e-mail címe:",
            "Az anya telefonszáma:", 
            "Az anya állandó lakhelye:",
            "Melyik óvodába járt a gyermek:",
            "Milyen jellegű osztályt választana (több lehetőség is választható):",
            "Választható tantárgy:", 
            "Napköziotthon:",
            "Iskolai étkeztetésre igényt tart:",
            "Van a gyermekének allergiája vagy más betegsége, melyről az iskolálak tudnia kell?:",
            "A szülők egy háztartásban élnek?",
            "Elsődleges kapcsolattartási személy vezetékneve és keresztneve (kit kereshetünk iskolai ügyekben):",
            "Elsődleges kapcsolattartási telefonszám (kit kereshetünk iskolai ügyekben):",
            "Elsődleges kapcsolattartási e-mail cím (kit kereshetünk iskolai ügyekben):", 
            "Az iskolalátogatásról szóló határozatot, illetve más levelezést az iskola kinek a nevére címezheti?:",
            "Bármilyen egyéb megjegyzés, amiről esetleg tudnunk kellene:"
        ]
        
        # Kezdjük a dokumentumot a preambulummal
        result = preamble + "\n"
        
        # Beiratkozási lapok generálása
        for i, data in enumerate(all_data, 1):
            # A sablon másolata az aktuális bejegyzéshez
            entry = entry_template
            
            # Változók listájának előkészítése
            variables = []
            
            # Mezők kitöltése a megfelelő sorrendben
            for field in field_order:
                if field in data and data[field]:
                    value = escape_latex(data[field])
                else:
                    value = ""
                variables.append(value)
            
            # Változók behelyettesítése
            for value in variables:
                # Csak az első előfordulást cseréljük le
                entry = entry.replace('!!!VARIABLE!!!', value, 1)
            
            # A sorszám behelyettesítése
            entry = entry.replace('BSSz: {0}', f'BSSz: {i}')
            
            # Hozzáadás az eredményhez
            result += entry
        
        # Dokumentum lezárása
        result += footer
        
        # Mentés
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(result)
            
        print(f"TeX fájl sikeresen mentve: {output_path}")
        
    except Exception as e:
        print(f"Hiba a TeX fájl készítése közben: {str(e)}")
        raise

def save_to_csv(all_data, output_path):
   """Menti az adatokat CSV fájlba."""
   try:
       # CSV fejléc (minden mező)
       fieldnames = list(all_data[0].keys())
       
       with codecs.open(output_path, 'w', encoding='utf-8-sig', errors='replace') as f:
           writer = csv.DictWriter(f, fieldnames=fieldnames)
           writer.writeheader()
           writer.writerows(all_data)
           
       print(f"CSV fájl sikeresen mentve: {output_path}")
   except Exception as e:
       print(f"Hiba a CSV mentése közben: {str(e)}")
       raise

def main():
   try:
       # Aktuális könyvtár
       current_dir = os.path.dirname(os.path.abspath(__file__))
       if not current_dir:  # Ha üres, akkor az aktuális munkakönyvtár
           current_dir = os.getcwd()
       
       print(f"Munkakönyvtár: {current_dir}")
       
       # Sablon és kimeneti fájlok
       template_path = os.path.join(current_dir, "adatlap.tex")  # A TeX sablon
       output_csv_path = os.path.join(current_dir, "beiratkozasok.csv")  # A kimeneti CSV
       output_tex_path = os.path.join(current_dir, "beiratkozasok.tex")  # A kimeneti TeX
       
       print(f"TeX sablon: {template_path}")
       print(f"CSV kimenet: {output_csv_path}")
       print(f"TeX kimenet: {output_tex_path}")
       
       # Az összes .eml fájl feldolgozása
       all_data = []
       
       for filename in os.listdir(current_dir):
           if filename.lower().endswith('.eml'):
               print(f"Feldolgozás: {filename}")
               eml_path = os.path.join(current_dir, filename)
               try:
                   data = process_eml_file(eml_path)
                   if data:
                       all_data.append(data)
               except Exception as e:
                   print(f"Hiba a '{filename}' fájl feldolgozása közben: {str(e)}")
       
       # Ha vannak adatok, mentjük őket
       if all_data:
           print(f"Összesen {len(all_data)} beiratkozási adat feldolgozva.")
           
           # CSV mentése
           save_to_csv(all_data, output_csv_path)
           
           # TeX fájl generálása a sablonból
           create_tex_file(all_data, output_tex_path, template_path)
       else:
           print("Nem sikerült adatokat kinyerni az .eml fájlokból.")
   
   except Exception as e:
       print(f"Kritikus hiba a program futása közben: {str(e)}")

if __name__ == "__main__":
   main()