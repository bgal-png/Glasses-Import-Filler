import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO

# ==========================================
# 🛑 VERSION CHECK 
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "4.2 - ZERO-STRIPPER ACTIVE"

st.title(f"🏭 Manufacturer Data Linker")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

# ==========================================
# 🎯 THE TARGET MAPPING (Your File)
# ==========================================
TARGET_MAPPING = {
    "Combination": "Combination (size on glasses)",
    "Barcode": "Barcode",
    "Glasses_type": "Glasses type ID: 13",
    "Manufacturer": "Manufacturer ID: 9", 
    "Glasses_size_temple_length": "Glasses size: temple length ID: 70",
    "Glasses_size_lens_height": "Glasses size: lens height ID: 71",
    "Glasses_size_lens_width": "Glasses size: lens width ID: 72",
    "Glasses_size_bridge": "Glasses size: bridge ID: 73",
    "Glasses_shape": "Glasses shape ID: 25",
    "Glasses_other_info": "Glasses other info ID: 49",
    "Glasses_frame_type": "Glasses frame type ID: 50",
    "Frame_Colour": "Frame Colour ID: 26",
    "Temple_Colour": "Temple Colour ID: 39",
    "Glasses_main_material": "Glasses main material ID: 53",
    "Glasses_lens_Colour": "Glasses lens Colour ID: 28",
    "Glasses_lens_material": "Glasses lens material ID: 35",
    "Glasses_lens_effect": "Glasses lens effect ID: 37",
    "Sunglasses_filter": "Sunglasses filter ID: 77",
    "Glasses_gendre": "Glasses gendre ID: 22",
    "Glasses_collection": "Glasses collection ID: 33",
    "SunGlasses_RX_lenses": "SunGlasses RX lenses ID:108",
    "Brand": "Brand ID:11",
    "Case_weight_g": "Case weight (g)",
    "Glasses_weight_g": "Glasses weight (g)",
    "Item_origin_country": "Item origin country",
    "Producing_company": "Producing company ID:146" 
}
# ==========================================
# 🔤 THE VALUE TRANSLATOR (Standardizing Data)
# ==========================================
# Format -> "Global_Column_Name": { "Their_Value": "Our_System_Value" }
# IMPORTANT: Put ALL variations from ALL manufacturers in the same list!
VALUE_TRANSLATOR = {
    "Glasses_type": {
        "": "",
        #safilo
        "SUNGLASSES FRAMES":"Sunglasses",
        "CLIP-ON":"Sunglasses",
        "CLIP-IN/ADAPTOR":"Sunglasses",
        "SUN + CLIP-IN": "Sunglasses",
        "VISOR/CUP": "Sunglasses",
        "BLUE BLOCK SUN": "Sunglasses",
        "OPT + CLIP-ON": "Frames",
        "OPTICAL FRAMES": "Frames",
        "SUNCOVER": "Frames",
        "BLUE BLOCK": "PC Glasses without power",
        "READERS + CLIP-ON": "PC Glasses without power",
        "READERS": "Reading glasses",
        #luxottica
        "Sluneční Brýle": "Sunglasses",
        "Dětské Sluneční Brýle": "Sunglasses",
        "Dětské Brýle": "Frames",
        "Brýle": "Frames",
        #kering
        "Optical Frame":"Frames",
        "Sunglass":"Sunglasses",
        "Sunglasses": "Sunglass",
        "Frames": "Optical Frame",
        #marcholin
        "Optical Frame":"Frames",
        "Sunglass":"Sunglasses",
        "Sunglasses": "Sunglass",
        "Frames": "Optical Frame",
    },
    "Glasses_shape": { #na shape až budou kompletní data od výrobců
    #safilo
    "OTHER SHAPE": "Extravagant",
    "ROUND GEOMETRICAL": "Round",
    "ROUND": "Round",
    "PILOT": "Pilot",
    "NAVIGATOR": "Pilot",
    "CAT EYE": "Cat Eye",
    "SQUARE": "Square",
    "SQUARE FLAT TOP": "Square",
    "SQUARE DOUBLE BRIDGE": "Square",
    "SQUARE GEOMETRICAL": "Square",
    "RECTANGULAR GEOMETRICAL": "Rectangular",
    "RECTANGULAR FLAT TOP": "Rectangular",
    "PANTHOS": "Panthos / Tea cup",
    "OVAL": "Oval / Elipse",
    "MASK": "Single lens",
    "BUTTERFLY": "Butterfly",
    "BUTTERFLY GEOMETRICAL": "Butterfly",
    "RECTANGULAR BROWLINE": "Browline",
    #luxottica - napsat dotaz na kategorie
    #kering
    "SHIELD": "Single lens",
    "MASK": "Single lens",
    "SQUARED": "Square",
    "PILOT/NAVIGATOR": "Pilot",
    "RECTANGULAR": "Rectangular",
    "ROUND": "Round",
    "PANTHOS": "Panthos / Tea cup",
    "OVAL": "Oval / Elipse",
    "BUTTERFLY": "Butterfly",
    #marcolin
    "SHIELD": "Single lens",
    "MASK": "Single lens",
    "SQUARED": "Square",
    "PILOT/NAVIGATOR": "Pilot",
    "RECTANGULAR": "Rectangular",
    "ROUND": "Round",
    "PANTHOS": "Panthos / Tea cup",
    "OVAL": "Oval / Elipse",
    "BUTTERFLY": "Butterfly",
    },
    "Glasses_other_info": {
        "": "",
        #safilo
        "SQUARE DOUBLE BRIDGE": "Double bridge",
        #luxottica
        #marcolin
        "Flex":"Flex",
        #kering
        "Flex":"Flex",
    },
    "Glasses_frame_type": {
        #safilo
        "FULL RIM": "Full rim",
        "RIMLESS": "Rimless",
        "HALF RIM": "Half rim",
        #luxottica
        "plná obruba": "Full rim",
        "bezobroučkové": "Rimless",
        "polo obruba": "Half rim",
        #marcolin
        "Full rim": "Full rim",
        "Rimless": "Rimless",
        "Half rim": "Half rim",
        #kering
        "Full rim": "Full rim",
        "Rimless": "Rimless",
        "Half rim": "Half rim",
    },
    "Frame_Colour": {
        "": "",
        #safilo
        "0": "",
        "BLACK": "Black",
        "WHITE": "White",
        "RED": "Red",
        "BLUE": "Blue",
        "AZURE": "Blue",
        "GREEN": "Green",
        "VIOLET": "Purple",
        "ORANGE": "Orange",
        "GREY": "Grey",
        "YELLOW": "Yellow",
        "BROWN": "Brown",
        "GOLD": "Gold",
        "SILVER": "Silver",
        "PINK": "Pink",
        "BEIGE": "Ivory",
        "IVORY": "Ivory",
        "HORN": "Ivory",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "CRYSTAL": "Havana",
        "BURGUNDY": "Burgundy",
        "TURQUOISE": "Turquoise",
        "RUTHENIUM": "Ruthenium",
        "ROSE GOLD": "Rose Gold",
        #luxottica
        "černá": "Black",
        "bílá": "White",
        "červená": "Red",
        "modrá": "Blue",
        "Zelená": "Green",
        "fialová": "Purple",
        "- Oranžové.": "Orange",
        "šedá": "Grey",
        "žlutá": "Yellow",
        "hnědá": "Brown",
        "zlato": "Gold",
        "stříbrná": "Silver",
        "růžová": "Pink",
        "béžová": "Ivory",
        "vícebarevné": "Multicolor",
        "transparentní": "Transparent",
        "žíhaná": "Havana",
        "burgundská": "Burgundy",
        "tyrkysová": "Turquoise",
        "měděná": "Ruthenium",
        "růžové zlato": "Rose Gold",
        #marcolin
        "BLACK": "Black",
        "WHITE": "White",
        "RED": "Red",
        "BLUE": "Blue",
        "LIGHT-BLUE": "Blue",
        "GREEN": "Green",
        "VIOLET": "Purple",
        "ORANGE": "Orange",
        "GREY": "Grey",
        "YELLOW": "Yellow",
        "BROWN": "Brown",
        "GOLD": "Gold",
        "COPPER": "Gold",
        "SILVER": "Silver",
        "PINK": "Pink",
        "FUCHSIA": "Pink",
        "BEIGE": "Ivory",
        "IVORY": "Ivory",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "NUDE": "Transparent",
        "HAVANA": "Havana",
        "CRYSTAL": "Havana",
        "BURGUNDY": "Burgundy",
        "LIGHT-BLUE": "Turquoise",
        "TURQUOISE": "Turquoise",
        "RUTHENIUM": "Ruthenium",
        "BRONZE": "Ruthenium",
        "ROSE GOLD": "Rose Gold",
        #kering
        "BLACK": "Black",
        "WHITE": "White",
        "RED": "Red",
        "BLUE": "Blue",
        "LIGHT-BLUE": "Blue",
        "GREEN": "Green",
        "VIOLET": "Purple",
        "ORANGE": "Orange",
        "GREY": "Grey",
        "YELLOW": "Yellow",
        "BROWN": "Brown",
        "GOLD": "Gold",
        "COPPER": "Gold",
        "SILVER": "Silver",
        "PINK": "Pink",
        "FUCHSIA": "Pink",
        "BEIGE": "Ivory",
        "IVORY": "Ivory",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "NUDE": "Transparent",
        "HAVANA": "Havana",
        "CRYSTAL": "Havana",
        "BURGUNDY": "Burgundy",
        "LIGHT-BLUE": "Turquoise",
        "TURQUOISE": "Turquoise",
        "RUTHENIUM": "Ruthenium",
        "BRONZE": "Ruthenium",
        "ROSE GOLD": "Rose Gold",
    },
    "Temple_Colour": {
        "": "",
        #safilo
        "0":"Havana",
        "BLACK": "Black",
        "WHITE": "White",
        "RED": "Red",
        "BLUE": "Blue",
        "AZURE": "Blue",
        "GREEN": "Green",
        "VIOLET": "Purple",
        "ORANGE": "Orange",
        "GREY": "Grey",
        "YELLOW": "Yellow",
        "BROWN": "Brown",
        "GOLD": "Gold",
        "SILVER": "Silver",
        "PINK": "Pink",
        "BEIGE": "Ivory",
        "IVORY": "Ivory",
        "HORN": "Ivory",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "CRYSTAL": "Havana",
        "BURGUNDY": "Burgundy",
        "TURQUOISE": "Turquoise",
        "RUTHENIUM": "Ruthenium",
        "ROSE GOLD": "Rose Gold",
        #luxottica
        "černá": "Black",
        "bílá": "White",
        "červená": "Red",
        "modrá": "Blue",
        "Zelená": "Green",
        "fialová": "Purple",
        "- Oranžové.": "Orange",
        "šedá": "Grey",
        "žlutá": "Yellow",
        "hnědá": "Brown",
        "zlato": "Gold",
        "stříbrná": "Silver",
        "růžová": "Pink",
        "béžová": "Ivory",
        "vícebarevné": "Multicolor",
        "transparentní": "Transparent",
        "žíhaná": "Havana",
        "burgundská": "Burgundy",
        "tyrkysová": "Turquoise",
        "měděná": "Ruthenium",
        "růžové zlato": "Rose Gold",
        #marcolin
        "SHINY":"",
        "NONE":"",
        "BLACK": "Black",
        "WHITE": "White",
        "RED": "Red",
        "BLUE": "Blue",
        "LIGHT-BLUE": "Blue",
        "GREEN": "Green",
        "VIOLET": "Purple",
        "ORANGE": "Orange",
        "GREY": "Grey",
        "YELLOW": "Yellow",
        "BROWN": "Brown",
        "GOLD": "Gold",
        "COPPER": "Gold",
        "SILVER": "Silver",
        "PINK": "Pink",
        "FUCHSIA": "Pink",
        "BEIGE": "Ivory",
        "IVORY": "Ivory",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "NUDE": "Transparent",
        "HAVANA": "Havana",
        "CRYSTAL": "Havana",
        "BURGUNDY": "Burgundy",
        "LIGHT-BLUE": "Turquoise",
        "TURQUOISE": "Turquoise",
        "RUTHENIUM": "Ruthenium",
        "BRONZE": "Ruthenium",
        "ROSE GOLD": "Rose Gold",
        #kering
        "NONE":"",
        "BLACK": "Black",
        "WHITE": "White",
        "RED": "Red",
        "BLUE": "Blue",
        "LIGHT-BLUE": "Blue",
        "GREEN": "Green",
        "VIOLET": "Purple",
        "ORANGE": "Orange",
        "GREY": "Grey",
        "YELLOW": "Yellow",
        "BROWN": "Brown",
        "GOLD": "Gold",
        "COPPER": "Gold",
        "SILVER": "Silver",
        "PINK": "Pink",
        "FUCHSIA": "Pink",
        "BEIGE": "Ivory",
        "IVORY": "Ivory",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "NUDE": "Transparent",
        "HAVANA": "Havana",
        "CRYSTAL": "Havana",
        "BURGUNDY": "Burgundy",
        "LIGHT-BLUE": "Turquoise",
        "TURQUOISE": "Turquoise",
        "RUTHENIUM": "Ruthenium",
        "BRONZE": "Ruthenium",
        "ROSE GOLD": "Rose Gold",
    },
    "Glasses_main_material": {
        #safilo
        "PLASTIC": "Plastic",
        "METAL": "Metal",
        "TITANIUM": "Titanium",
        "WOOD": "Wood",
        #luxottica
        "Nylon": "Plastic",
        "acetát": "Plastic",
        "polyéteréterketon (PEEK)": "Plastic",
        "O_matter": "Plastic",
        "jiný materiál": "Plastic",
        "propionát": "Plastic",
        "rohovina": "Plastic",
        "Tvarovaný acetát": "Plastic",
        "Karbonové vlákno": "Plastic",
        "Ocel": "Metal",
        "C_5": "Metal",
        "hliník": "Metal",
        "Kov s paměťovým efektem": "Metal",
        "Titan": "Titanium",
        "Dřevo": "Wood",
        #marcolin
        "INJECTION": "Plastic",
        "RECYCLED ACETATE": "Plastic",
        "BIO INJECTION": "Plastic",
        "NYLON": "Plastic",
        "BIO NYLON": "Plastic",
        "BIO INJECTION RILSAN": "Plastic",
        "BIO ACETATE": "Plastic",
        "HORN": "Plastic",
        "RECYCLED INJECTION": "Plastic",
        "OPTYL": "Plastic",
        "MEMORY METAL": "Plastic",
        "RECYCLED INJECTED ACETATE": "Plastic",
        "RUBBER": "Plastic",
        "POLYCARBONATE": "Plastic",
        "METAL": "Metal",
        "ALLUMINIUM": "Metal",
        "GOLD": "Metal",
        "COPPER": "Metal",
        "STAINLESS STEEL": "Metal",
        "TITANIUM": "Titanium",
        "WOOD": "Wood",
        #kering
        "INJECTION": "Plastic",
        "RECYCLED ACETATE": "Plastic",
        "BIO INJECTION": "Plastic",
        "NYLON": "Plastic",
        "BIO NYLON": "Plastic",
        "BIO INJECTION RILSAN": "Plastic",
        "BIO ACETATE": "Plastic",
        "HORN": "Plastic",
        "RECYCLED INJECTION": "Plastic",
        "OPTYL": "Plastic",
        "MEMORY METAL": "Plastic",
        "RECYCLED INJECTED ACETATE": "Plastic",
        "RUBBER": "Plastic",
        "POLYCARBONATE": "Plastic",
        "METAL": "Metal",
        "ALLUMINIUM": "Metal",
        "GOLD": "Metal",
        "COPPER": "Metal",
        "STAINLESS STEEL": "Metal",
        "TITANIUM": "Titanium",
        "WOOD": "Wood",

    },
    "Glasses_lens_Colour": {
        "": "",
        #safilo
        "0": "",
        "BLACK": "Black",
        "RED": "Red",
        "BLUE": "Blue",
        "GREEN": "Green",
        "GOLD": "Gold",
        "SILVER": "Silver",
        "GREY": "Grey",
        "ORANGE": "Orange",
        "YELLOW": "Yellow",
        "VIOLET": "Purple",
        "BROWN": "Brown",
        "PINK": "Pink",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "BURGUNDY": "Burgundy",
        #luxottica
        "černá": "Black",
        "Black": "Black",
        "červená": "Red",
        "Red": "Red",
        "modrá": "Blue",
        "Blue": "Blue",
        "zelená": "Green",
        "Green": "Green",
        "zlatá": "Gold",
        "Gold": "Gold",
        "střírbrná": "Silver",
        "Silver": "Silver",
        "šedá": "Grey",
        "Grey": "Grey",
        "oranžová": "Orange",
        "Orange": "Orange",
        "žlutá": "Yellow",
        "Yellow": "Yellow",
        "fialová": "Purple",
        "Purple": "Purple",
        "hnědá": "Brown",
        "Brown": "Brown",
        "růžová": "Pink",
        "Pink": "Pink",
        "vícebarvné": "Multicolor",
        "Multicolor": "Multicolor",
        "průhledné": "Transparent",
        "Transparent": "Transparent",
        "burgundská": "Burgundy",
        "tmavě zelená": "Green",
    "tmavě šedá": "Grey",
    "zelená, zrcadlové, zelená": "Green",
    "polarizační, tmavě šedá": "Grey",
    "tmavě šedá, polarizační": "Grey",
    "Red Hiper 8": "Red",
    "žlutá": "Yellow",
    "Šedá Zrcadlová Oranžová/Žlutá": "Multicolor",
    "Tmavá Šedá Zrcadlová Voda Polar": "Grey",
    "světle šedá, zrcadlové, modrá": "Grey|Blue",
    "polarizační, šedá": "Grey",
    "šedá, zrcadlové, stříbrná": "Grey|Silver",
    "modrá, zrcadlová, modrá": "Blue",
    "tmavě zelená, polarizační": "Green",
    "tmavě hnědá": "Brown",
    "Černá Kouřová": "Black|Grey",
    "modrá, gradální, šedá": "Blue|Grey",
    "Tmavě fialová": "Purple",
    "Polarizovaná Tmavá Šedá": "Grey",
    "světle šedá, zrcadlové, černá": "Grey|Black",
    "- Oranžové.": "Orange",
    "šedá, zrcadlové, fialová": "Grey|Purple",
    "Polar Tmavá Modrá": "Blue",
    "modrá, zrcadlové": "Blue",
    "modrá": "Blue",
    "Grey Mirror Silver 80 Polarized": "Grey|Silver",
    "šedá, černá": "Grey|Black",
    "šedá, zrcadlové, černá": "Grey|Black",
    "tmavě modrá": "Blue",
    "Polarizovaná Tmavá Šedá Zrcadlová Voda": "Grey",
    "šedá, polarizační": "Grey",
    "světle šedá, zrcadlové, stříbrná 80": "Grey|Silver",
    "Grey Mirror Orange/Yellow ": "Multicolor",
    "Zelená Polarizovaná": "Green",
    "Hnědá Polarizovaná": "Brown",
    "Tmavá Modrá Polar": "Blue",
    "Zrcadlová Žlutá Zlatá": "Yellow|Gold",
    "světle azurová, zrcadlové, čírka": "Blue|Transparent",
    "Tmavá Zelená Zrcadlová Petrolejová": "Green|Blue",
    "stříbrná": "Silver",
    "světle zelená": "Green",
    "Safírová Modrá Zrcadlová": "Blue",
    "Zelená": "Green",
    "Fifty Blue/Orange": "Blue|Orange",
    "Fifty Black/Yellow": "Black|Yellow",
    "Fifty Brown/Blue": "Brown|Blue",
    "Red/Dark Grey": "Red|Grey",
    "Yellow/Dark Grey": "Yellow|Grey",
    "gradální, hnědá": "Brown",
    "Růže Gradient Šedá Zrcadlová Modrá": "Multicolor",
    "modrá/zelená": "Blue|Green",
    "tmavě šedá, zrcadlové, červená/žlutá": "Multicolor",
    "zrcadlově světle šedo-stříbrná": "Grey|Silver",
    "světle zelená, zrcadlové, petrolejová": "Green|Blue",
    "čiré": "Transparent",
    "Grey Mirror Silver 80 Polar": "Grey|Silver",
    "Světlá Zelená Zrcadlová Modrá": "Green|Blue",
    "58 – POLARIZAČNÍ ZELENÁ": "Green",
    "polarizační, tmavě zelená": "Green",
    "tmavě modrá, zrcadlové, modrá": "Blue",
    "tmavě hnědá, polarizační": "Brown",
    "hnědá, zrcadlové, zlato": "Brown|Gold",
    "Červeno/žlutá zrcadlová": "Red|Yellow",
    "Mirror Water Polar": "Blue",
    "polarizační, stříbrná, zrcadlové": "Silver",
    "polarizační, hnědá": "Brown",
    "transparentní": "Transparent",
    "Odstín Recycled Demo": "Transparent",
    "hnědá": "Brown",
    "azurová": "Blue",
    "Fialová Zrcadlová Flash Stříbrná": "Purple|Silver",
    "světle hnědá": "Brown",
    "světle šedá": "Grey",
    "šedá": "Grey",
    "šedá gradální": "Grey",
    "zelená, gradální": "Green",
    "Gradient Zelená": "Green",
    "Photocromatic Brown": "Brown",
    "čiré, gradální, hnědá": "Transparent|Brown",
    "gradální, hnědá, zrcadlové, stříbrná": "Brown|Silver",
    "Polarizační přechodová šedá": "Grey",
    "Hnědá Zrcadlová Stříbrná Gradient": "Brown|Silver",
    "Foto Hnědá": "Brown",
    "Jasná Zrcadlová Skutečná Růže Zlatá": "Pink|Gold",
    "Clear Mirror Real Yellow Gold": "Multicolor",
    "světle modrá": "Blue",
    "Foto Šedá": "Grey",
    "Light Grey Tampo Flower Silver": "Grey|Silver",
    "Light Violet Tampo Flower Silver": "Purple|Silver",
    "Čirá Gradient Světlá Hnědá": "Transparent|Brown",
    "čiré, gradální, světle modrá": "Transparent|Blue",
    "Šedá gradientní zelená": "Grey|Green",
    "světle modrá stříbrná, zrcadlové": "Blue|Silver",
    "Olivově zelená": "Green",
    "Exkluzivní modrá": "Blue",
    "Exkluzivní zelená": "Green",
    "Exkluzivní hnědá": "Brown",
    "Světlá Purpurová Hnědá": "Purple|Brown",
    "hnědá, gradální": "Brown",
    "světle hnědá, gradální, šedá": "Brown|Grey",
    "gradální, šedá": "Grey",
    "Šedá/Zelená": "Grey|Green",
    "Hnědé Zrcadlová Gradient Šedá": "Brown|Grey",
    "růžová, gradální, tmavě fialová": "Pink|Purple",
    "Zelená Zrcadlová Stříbrná": "Green|Silver",
    "modrá, gradální": "Blue",
    "světle modrá, gradální, šedá": "Blue|Grey",
    "Čirá Gradient Hnědá Foto": "Transparent|Brown",
    "Čirá Gradient Šedá": "Transparent|Grey",
    "polarizační, zelená": "Green",
    "Clear Blue-Violet Light Filter": "Transparent|Blue",
    "světle šedá, gradální, žlutá": "Grey|Yellow",
    "Čirá fotochromatická až tmavý křemen": "Transparent|Grey",
    "Růžová Gradient Šedá": "Pink|Grey",
    "červená": "Red",
    "Fotochromatická šedá": "Grey",
    "šedá/modrá": "Grey|Blue",
    "Photochromatic Brown": "Brown",
    "čiré, gradální, tmavě zelená": "Transparent|Green",
    "Žlutá Gradient Světle Hnědá": "Yellow|Brown",
    "Grey Vintage Blue": "Grey|Blue",
    "Oranžová zrcadlová stříbrná": "Orange|Silver",
    "růžová, gradální, fialová": "Pink|Purple",
    "Šedá Gradient Hnědá": "Grey|Brown",
    "Grey Blue Black Gradient": "Multicolor",
    "Dark Blue Polarized": "Blue",
    "Zrcadlová Černá": "Black",
    "57 – POLARIZAČNÍ HNĚDÁ": "Brown",
    "polarizační, gradální, hnědá": "Brown",
    "Šedá Zrcadlová Stříbrná Polar": "Grey|Silver",
    "Čirá Gradient Hnědá Zrcadlová Stříbrná": "Multicolor",
    "Oranžová Gradient Hnědá": "Orange|Brown",
    "Světle Modrá Zrcadlová Stříbrná Uvnitř": "Blue|Silver",
    "kouř": "Grey",
    "Zrcadlová Stříbrná Polar": "Silver",
    "šedá, gradální, modrá": "Grey|Blue",
    "gradální, modrá": "Blue",
    "čiré, gradální, modrá": "Transparent|Blue",
    "Zrcadlová Modrá": "Blue",
    "tmavě fialová, zrcadlové, červená": "Purple|Red",
    "Fialová Zrcadlová Stříbrná Gradient": "Purple|Silver",
    "Azurová Gradient Tmavá Modrá": "Blue",
    "Zrcadlová Zelená": "Green",
    "Polarizační hnědo-zlaté zrcadlová": "Brown|Gold",
    "Růžový Gradient Hnědá": "Pink|Brown",
    "Světlá Šedá Flash Stříbrná": "Grey|Silver",
    "hnědá, gradální, tmavě hnědá": "Brown",
    "Dark Grey Voodoo": "Grey",
    "Dark Brown Voodoo": "Brown",
    "Green Voodoo": "Green",
    "bronzová": "Orange",
    "Světlá Zelená Zrcadlová Zelená": "Green",
    "Zrcadlově Oranžová": "Orange",
    "Hnědá zrcadlová – vnitřní stříbrná": "Brown|Silver",
    "Fialová Zrcadlová Uvnitř Stříbrná": "Purple|Silver",
    "Tmavá Oranžová Zrcadlová Zlatá": "Orange|Gold",
    "Žlutá Gradient Hnědá": "Yellow|Brown",
    "zrcadlově tmavě šedo-stříbrná": "Grey|Silver",
    "Brown Grad Purple Grad Black": "Multicolor",
    "žlutá/hnědá": "Yellow|Brown",
    "růžová, gradální, tmavě šedá": "Pink|Grey",
    "čiré, s filtrem modrého světla": "Transparent|Blue",
    "Polar Hnědá Gradient": "Brown",
    "Gradient Růžová": "Pink",
    "Grey Border Red/Beige": "Multicolor",
    "Čirá Gradient Růžová": "Transparent|Pink",
    "Čirá Gradient Červená": "Transparent|Red",
    "Světlá Hnědá Zrcadlová Zlatá": "Brown|Gold",
    "Dark Brown Tampo Gold": "Brown|Gold",
    "Petrol Green Mirror Silver": "Green|Silver",
    "fialová": "Purple",
    "Purple Polarized": "Purple",
    "Tmavá Růžová": "Pink",
    "polarizační, šedá, gradální": "Grey",
    "Šedá Tampo Borůvková Stříbrná/Zlatá": "Multicolor",
    "Light Grey Tampo Silver": "Grey|Silver",
    "Šedá Tampo Borůvková Stříbrná/Zlatá2": "Multicolor",
    "polarizační, hnědá, gradální": "Brown",
    "Světlá Šedá Gradient Tmavě Modrá": "Grey|Blue",
    "Růže Gradient Šedá": "Pink|Grey",
    "Tmavá Šedá": "Grey",
    "Hnědá Polar": "Brown",
    "Světlá Hnědá Zrcadlová Flash Zlatá": "Brown|Gold",
    "Čirá Zrcadlová Stříbrná 80": "Transparent|Silver",
    "Marrone": "Brown",
    "Světle šedá stříbrná, odstín Tampo Mirror Silver": "Grey|Silver",
    "tmavě purpurová": "Purple",
    "Rosa Specchio Argento": "Pink|Silver",
    "Gradient Šedá Polar": "Grey",
    "Dark Gery": "Grey",
    "Světlá Fialová Flash Stříbrná": "Purple|Silver",
    "Grigio Sfumato Blu": "Grey|Blue",
    "Grigio Specchio Argento": "Grey|Silver",
    "Verde Specchio Petrol": "Green|Blue",
    "Blu Specchio": "Blue",
    "Marrone Specchio Oro": "Brown|Gold",
    "Jednobarevná Hnědá": "Brown",
    "Fialová Gradient Tmavá Šedá": "Purple|Grey",
    "Světle Šedá Gradient Černá": "Grey|Black",
    "růžová": "Pink",
    "Polarizovaná Černá": "Black",
    "Grad Light Brown Mirror Gold": "Brown|Gold",
    "Zrcadlový odstín v barvě slonoviny": "Yellow",
    "Azurová/Růžová/Hnědá Gradient": "Multicolor",
    "Světlá Růžová Gradient Růžová": "Pink",
"růžová/hnědá": "Pink|Brown",
"Pink/Blue/Brown": "Multicolor",
"Pink Gradient Brown Mirror": "Pink|Brown",
"Šedá Zrcadlová Stříbrná Gradient": "Grey|Silver",
"Čiré zrcadlo – žluté zlato": "Transparent|Gold",
"Čirá Zrcadlová Stříbrná": "Transparent|Silver",
"plná modrá": "Blue",
"Světlá A Tmavě Hnědá Gradient": "Brown",
"Šedá Gradient Černá": "Grey|Black",
"Růžová zrcadlová – vnitřní stříbrná": "Pink|Silver",
"Purpe Brown": "Purple|Brown",
"Oranžová Gradient Světlá Zelená": "Orange|Green",
"Čirá Gradient Tmavá Brandy": "Transparent|Brown",
"Světlá Žlutá Gradient Světlá Hnědá": "Yellow|Brown",
"světle fialová": "Purple",
"Světlá Žlutá Gradient Šedá": "Yellow|Grey",
"Čirá fotochromatická až šedý křemen": "Transparent|Grey",
"Modrozelená": "Blue|Green",
"Světlá Zelená Gradient Tmavá Hnědá": "Green|Brown",
"Čirá / hnědá / stříbrná": "Multicolor",
"Růžová Gradient Červená": "Pink|Red",
"Žlutá Gradient Šedá": "Yellow|Grey",
"Orange/Violet": "Orange|Purple",
"Světlá Hnědá Gradient Zelená": "Brown|Green",
"Tmavá Šedá Tampo D&G Stříbrná": "Grey|Silver",
"Light Pink Mirror Silver": "Pink|Silver",
"Dark Grey Ar Blue": "Grey|Blue",
"purpurová": "Purple",
"Růžová Zrcadlová Růžová Zlatá": "Pink|Gold",
"Tmavá Fialová Zrcadlová Stříbrná": "Purple|Silver",
"Žlutá Zrcadlová Červená": "Yellow|Red",
"Jednobarevná Modrá": "Blue",
"Růžová Zrcadlová Růžová": "Pink",
"Oranžová Zrcadlová Růžová": "Orange|Pink",
"Dark Grey Mirror Pink": "Grey|Pink",
"Růžová Zrcadlová Bílá": "Pink",
"Oranžová Zlatá Zrcadlová": "Orange|Gold",
"Tmavá Šedá Zrcadlová Modrá": "Grey|Blue",
"Modrá Zrcadlová Bílá": "Blue",
"Světle fialová, odstín zrcadlově stříbné, vnitřek": "Purple|Silver",
"Světle oranžová, zrcadlově stříbrná": "Orange|Silver",
"Polar Šedá Zrcadlová Stříbrná": "Grey|Silver",
"Černočervená zdrcalová": "Black|Red",
"světle hnědá, zrcadlové, tmavě zlatá": "Brown|Gold",
"Kouřově černá zrcadlová": "Grey|Black",
"světle růžová": "Pink",
"tmavá bronzová": "Orange",
"Hnědá Zrcadlová Růže Zlatá": "Multicolor",
"Přechodová hnědo-šedá": "Brown|Grey",
"Zelená Gradient Zelená": "Green",
"Brown Gradient Green": "Brown|Green",
"Přechodová čirá žlutá": "Transparent|Yellow",
"Gradient Blue Polarized": "Blue",
"Gradient Hnědá Polar": "Brown",
"Gradient Kouřová": "Grey",
"Tmavá Šedá Zrcadlová Červená": "Grey|Red",
"lila": "Purple",
"pruhovaná modrá": "Blue",
"Leopardí": "Multicolor",
"Gradient Fialová": "Purple",
"Grigio Polar": "Grey",
"Tmavě Šedá Ar Modrá Vnější": "Grey|Blue",
"Brown Grad Grey Mirror Silver": "Multicolor",
"purpurová hnědá": "Purple|Brown",
"Gradient Žlutá": "Yellow",
"Čirá Gradient Okrová": "Transparent|Brown",
"Hnědý gradient, světle zelená": "Brown|Green",
"Světle Oranžová": "Orange",
"světle azurová": "Blue",
"šedá, gradální, polarizační": "Grey",
"čiré, gradální, tmavě fialová": "Transparent|Purple",
"hnědá, gradální, šedá": "Brown|Grey",
"Čirá Gradient Zelená": "Transparent|Green",
"Polarizační modrá": "Blue",
"hnědá, gradální, polarizační": "Brown",
"Dark Green Polarized": "Green",
"Polar Purple Mirror Gold Rose": "Multicolor",
"Hnědá Gradient Fialová": "Brown|Purple",
"Clear Gradient Orange": "Transparent|Orange",
"Čiré gradientní fuchsiové stříbrné zrcadlo": "Multicolor",
"Azurová Gradient Modrá": "Blue",
"růžová gradální": "Pink",
"Neutrální Gradient Světle Šedá": "Grey",
"hnědá, gradální, tmavě fialová": "Brown|Purple",
"Gradient Šedá": "Grey",
"hnědá, gradální, růžová": "Brown|Pink",
"Azurová Flash Stříbrná": "Blue|Silver",
"Tmavá Šedá Třpytivá": "Grey",
"Růžová Třpytivá": "Pink",
"růžová, gradální": "Pink",
"Světlá Žlutá": "Yellow",
"polarizační, gradální, šedá": "Grey",
"Světlá Azurová Stříbrná Zrcadlová": "Blue|Silver",
"fialová, zrcadlové": "Purple",
"Hnědá Gradient Šedá Zrcadlová Stříbrná": "Multicolor",
"Polarizační přechodová hnědá": "Brown",
"Čirá Gradient Modrá Zrcadlová Stříbrná": "Multicolor",
"Gradient Fuxia Mirror Silver": "Pink|Silver",
"Fialové Vnitřní Zrcadlová Stříbrná": "Purple|Silver",
"Drak Grey": "Grey",
"Hnědá Gradient Polar": "Brown",
"Grigio Scuro": "Grey",
"Polar Šedá": "Grey",
"Fialově hnědá zrcadlově stříbrná": "Multicolor",
"Polarizovaná gradientní šedá": "Grey",
"Šedá gradientní růžová": "Grey|Pink",
"Zelená zrcadlová černá": "Green|Black",
"Gradient Fialová Zrcadlová Stříbrná": "Purple|Silver",
"Polar Tmavá Šedá": "Grey",
"Světle Modrá Gradient Tmavě Modrá": "Blue",
"Fialová Zrcadlová Fialová": "Purple",
"Fialová Zrcadlová Růžová": "Purple|Pink",
"Světlý Růžový Odstín": "Pink",
"Tmavá Šedá Gradient": "Grey",
"Bordová Gradient": "Burgundy",
"Hnědá Růžová Gradient": "Brown|Pink",
"stříbrná, zrcadlové": "Silver",
"Růžová Zlatá Zrcadlová": "Pink|Gold",
"Clear Fifty Light Brown Gradie": "Transparent|Brown",
"Růžová Zlatá Gradient Zrcadlová": "Pink|Gold",
"zlatá, zrcadlové": "Gold",
"Jantarová Gradient": "Orange",
"Námořnická Gradient": "Blue",
"Stříbrná Kaki Flash Gradient": "Silver|Green",
"Hnědá Zrcadlová Gradient": "Brown",
"Rose Gold Gradient": "Pink|Gold",
"růžová, gradální, zrcadlové": "Pink",
"Tmavá Šedá Jednobarevná": "Grey",
"Stříbrný blok s opakováním loga MK": "Silver",
"Bordová-Hnědá": "Burgundy|Brown",
"Zelená Jednobarevná": "Green",
"Hnědá Jednobarevná": "Brown",
"Šedá Zrcadlová Jednobarevná": "Grey",
"hnědočerná": "Brown|Black",
"Růžová Černá": "Pink|Black",
"Šedá Černá": "Grey|Black",
"hnědá, zrcadlové": "Brown",
"Tmavá Šedá Zrcadlová": "Grey",
"světle šedá, gradální": "Grey",
"Olivová Zelená": "Green",
"Světlá Hnědá Jednobarevná": "Brown",
"Růže Jednobarevná": "Pink",
"Modrá Jednobarevná": "Blue",
"Šedá Jednobarevná": "Grey",
"šedá gradální, polarizační": "Grey",
"Chocolate Solid": "Brown",
"Plná velbloudí": "Brown",
"Jednobarevný muškátový oříšek": "Brown",
"zelená, polarizační": "Green",
"tmavě hnědá, gradální": "Brown",
"Khaki stříbrný gradient": "Green|Silver",
"Černá Zrcadlová": "Black",
"Stříbrná Flash Gradient": "Silver",
"Olive Mirror": "Green",
"modrá, polarizační": "Blue",
"Růžová Gradient Zrcadlově Růžová": "Pink",
"vínová": "Burgundy",
"kouřová, gradální": "Grey",
"Hroznová Jednobarevná": "Purple",
"Wisteria": "Purple",
"olivová, gradální": "Green",
"Plum Blue": "Purple",
"hnědá/stříbrná, zrcadlové": "Brown|Silver",
"Stříbrná Šedá Zrcadlová Gradient": "Silver|Grey",
"Bordovo-Hnědá Gradient": "Burgundy|Brown",
"Růžová Zlatá Polar": "Pink|Gold",
"Šedá Oranžová Gradient": "Grey|Orange",
"Modrá Šedá Gradient": "Blue|Grey",
"Purple Amber Gradient": "Purple|Orange",
"Hnědá Jednobarevná Polarizovaná": "Brown",
"Brown Sunset Polarized": "Brown",
"Světlá Hnědá Gradient Tmavá Hnědá": "Brown",
"Bordová Gradient Polar": "Burgundy",
"Šedá Modrá Gradient": "Grey|Blue",
"Ash Gradient": "Grey",
"Jantarová Jednobarevná": "Orange",
"světle hnědá, gradální": "Brown",
"Modrá Zrcadlová Gradient": "Blue",
"Tmavá Hnědá Jednobarevná Polar": "Brown",
"Tmavá Hnědá Jednobarevná": "Brown",
"Ash Solid": "Grey",
"Deep Red Solid": "Red",
"Light Iris Solid": "Purple",
"Green Solid Photochromic": "Green",
"Morušová Jednobarevná": "Purple",
"Hnědá Modrá Gradient": "Brown|Blue",
"Modrá Šedá Jednobarevná": "Blue|Grey",
"Růžová Gradient": "Pink",
"Tmavá Šedá Jednobarevná Polar": "Grey",
"Růžová Zlatá Zrcadlová Polar": "Pink|Gold",
"Hnědá Gradient Zrcadlová Červená": "Brown|Red",
"Purpurová Modrá Gradient": "Purple|Blue",
"zelená, gradální, zrcadlové": "Green",
"Soft Pink Gradient": "Pink",
"purpurová, gradální": "Purple",
"Růžová Gradient Růžová Zrcadlová Růžová": "Pink",
"Růžová Gradient Tmavě Hnědá": "Pink|Brown",
"Rose Flash": "Pink",
"Rose Mirror": "Pink",
"Hnědá Jednobarevná Polar": "Brown",
"Hnědá Stříbrná Zrcadlová": "Brown|Silver",
"Šedá Gradient Modrá Zrcadlová": "Grey|Blue",
"broskvová": "Pink",
"Šedá Čirá Gradient": "Grey|Transparent",
"hnědorůžová": "Brown|Pink",
"hnědomodrá": "Brown|Blue",
"jantarová": "Orange",
"olivová": "Green",
"Purpurová Zrcadlová": "Purple",
"růžové zlato": "Pink|Gold",
"Šedá Růže Gradient": "Grey|Pink",
"Hnědá zrcadlová stříbrná vnitřní fialová": "Multicolor",
"bordó": "Burgundy",
"Zelená Zrcadlová Uvnitř Stříbrná": "Green|Silver",
"čiré, stříbrná, gradální": "Transparent|Silver",
"Gradient Modrá Flash Stříbrná": "Blue|Silver",
"bledě zlatá": "Gold",
"Tmavě růžová zrcadlová – vnitřní stříbrná": "Pink|Silver",
"Ivory Mirror": "Yellow",
"šedá, gradální, zrcadlové": "Grey",
"Tmavě růžová stříbrná": "Pink|Silver",
"Hnědá Stříbrná": "Brown|Silver",
"Hnědá Zrcadlová Oranžová": "Brown|Orange",
"Yellow Gradient Silver": "Yellow|Silver",
"Gradient Modrá Zrcadlová Stříbrná": "Blue|Silver",
"Bronze Mirrow Gradient": "Orange",
"Růžová infračervená": "Pink",
"Čirá Gradient Tmavá Šedá": "Transparent|Grey",
"Heřmánková gradientní": "Yellow",
"Černá/Šedá": "Black|Grey",
"Filtr modrofialového světla": "Blue",
"Růžová Zrcadlová Stříbrná Gradient": "Pink|Silver",
"oranžová, zrcadlové, stříbrná, gradální": "Orange|Silver",
"Světlá Hnědá Zrcadlová Bronzová": "Brown|Orange",
"Prizm, černá, polarizační": "Black",
"Prizm, wolfram": "Brown",
"Prizm, šedá": "Grey",
"Prizm, nefritová": "Green",
"Prizm, safírová": "Blue",
"Prizm, fialová": "Purple",
"rubínová Prizm": "Red",
"Prizm 24K": "Gold",
"černá Prizm": "Black",
"Prizm Road, nefritová": "Green",
"Prizm, bronzová": "Orange",
"Prizm, 24 karátů": "Gold",
"Prizm Low, světlá": "Transparent",
"Prizm Deep Water, polarizační": "Blue",
"Prizm, rubínová, polarizační": "Red",
"Prizm, safírová, polarizační": "Blue",
"Prizm Trail Torch": "Red",
"černá, Iridium, polarizační": "Black",
"Prizm, wolfram polarizační": "Brown",
"Prizm, růžové zlato, polarizační": "Pink|Gold",
"Prizm Grey Gradient": "Grey",
"Prizm Brown Gradient": "Brown",
"Prizm, růžové zlato": "Pink|Gold",
"nefritová, Iridium": "Green",
"G40 černá, gradální": "Black",
"Prizm slate": "Grey",
"Prizm Road, černá": "Black",
"Prizm, broskvová": "Pink",
"Prizm, šedá, polarizační": "Grey",
"černá, Iridium": "Black",
"Prizm, 24 karátů, polarizační": "Gold",
"Prizm, nefritová, polarizační": "Green",
"pozitivní červení, Iridium": "Red",
"Prizm Shallow Water, polarizační": "Blue",
"čiré až černé, Iridium, fotochromatické": "Transparent|Black",
"chromovaná, Iridium": "Silver",
"růžová, gradální, polarizační": "Pink",
"Prizm Snow, safírová": "Blue",
"teplá šedá": "Grey",
"24K Iridium": "Gold",
"Prizm, bronzová, polarizační": "Orange",
"Prizm, fialová, polarizační": "Purple",
"ohnivá, Iridium": "Orange",
"Prizm Snow, černá, Iridium": "Black",
"wolfram, Iridium, polarizační": "Brown",
"safírová, Iridium": "Blue",
"Prizm Berry": "Purple",
"Prizm, indigo": "Blue",
"Prizm Torch Iridium": "Red",
"Prizm sage – zlaté iridium": "Green|Gold",
"Prizm snow – argon iridium": "Blue",
"Prizm, čiré": "Transparent",
"ohnivá, Iridium, polarizační": "Orange",
"24K Iridium Polarized": "Gold",
"fialová Iridium": "Purple",
"fialová Iridium, polarizační": "Purple",
"chromová, Iridium, polarizační": "Silver",
"Iridium, ledová": "Blue",
"Vr28 černá, Iridium, polarizační": "Black",
"rubínová, Iridium, polarizační": "Red",
"rubínová, Iridium": "Red",
"nefritová, Iridium, polarizační": "Green",
"safírová, Iridium, polarizační": "Blue",
"Prizm Snow, safírová Iridium": "Blue",
"polarizační, Torch Iridium": "Red",
"Prizm road – černé iridium": "Black",
"Prizm Snow Torch Iridium": "Red",
"Gradientní bronzová, odstín Prizm": "Orange",
"Prizm road – jadeitové iridium": "Green",
"Fotochromatická čirá černá iridium": "Transparent|Black",
"Zelená Láhev": "Green",
"Polarizovaná Tmavá Zelená": "Green",
"Šedá Modrá": "Grey|Blue",
"Mirror Light Grey": "Grey",
"oranžová, zrcadlové, červená": "Orange|Red",
"Gradient Černá": "Black",
"Starožitná Zelená": "Green",
"Oranžově-zlatá zrcadlová": "Orange|Gold",
"zrcadlové, červená": "Red",
"lesklá černá": "Black",
"bronzová hnědá, zrcadlové": "Orange|Brown",
"Dark Polar Brown": "Brown",
"Gradientní tmavá hnědá": "Brown",
"šedá, zrcadlové, modrá": "Grey|Blue",
"oranžová, zrcadlové": "Orange",
"Zrcadlová Šedá": "Grey",
"Světle Šedá Gradient Tmavě Šedá": "Grey",
"Hnědá Gradient Hnědá - Polar": "Brown",
"Žlutá Hnědá": "Yellow|Brown",
"polarizační, černá": "Black",
"Modrá Gradient Polarizovaná": "Blue",
"Photochromic Clear to Grey with Blue Light Filter": "Multicolor",
"Light Blue Gradient Dark Blue Polar": "Blue",
"modrý křišťál": "Blue",
"Polar Světle Modrá Gradient": "Blue",
"Polarizovaná tmavě fialová": "Purple",
"Modrá Světlá": "Blue",
"Transitions Signature Gen8 - Grey": "Grey",
"Transitions 8 Green": "Green",
"Transitions 8 Grey": "Grey",
"Polar Gradient Dark Blue": "Blue",
"Transitions Clear To Green": "Transparent|Green",
"Zelené Světlá": "Green",
"polarizační, gradální, modrá": "Blue",
"Demo Lens Blue Wash": "Transparent|Blue",
"Transitions Signature Gen8 - Sapphire": "Blue",
"Modrá Gradient Tmavá Modrá Polar": "Blue",
"Dark Polar": "Black",
"černé": "Black",
"Transitions 8 Brown": "Brown",
"Light Blue Gradient Dark Polarized": "Blue",
"Gradient Blue Polar": "Blue",
"Transitions 8 Sapphire": "Blue",
"Tmavě šedý gradient": "Grey",
"Blue Gradient Dark Blue Polarized": "Blue",
"Polar Black Grey": "Black|Grey",
"Transitions Dark Grey": "Grey",
"Transitions Clear To Sapphire": "Transparent|Blue",
"Modrá Starožitná": "Blue",
"Hnědá Gradient Hnědá": "Brown",
"Polar Černá + AR": "Black",
"Tmavá Fialová Polar + AR": "Purple",
"Žlutá Zrcadlová Zlatá": "Yellow|Gold",
"Žlutá Zlatá Zrcadlová": "Yellow|Gold",
"Graphite Sfumato Cristallo": "Grey",
"Grafite Cristallo": "Grey",
"Grey Gradient Vintage": "Grey",
"čiré, modrá": "Transparent|Blue",
"Foto Čirá Hnědá": "Transparent|Brown",
"Tmavá Fialová Zrcadlová Stříbrná Uvnitř": "Purple|Silver",
"Platinová": "Silver",
"Žlutá Oranžová": "Yellow|Orange",
"Yellow/Dark Blue": "Yellow|Blue",
"Terra Bio": "Brown",
"Petrolejová Zelená": "Green",
"Fucsia Crystal": "Pink",
"modrá, zrcadlové, stříbrná 80": "Blue|Silver",
"Světlá Hnědá Gradient Světlá Šedá": "Brown|Grey",
"Iris Sfumato Sole": "Purple",
"Žlutá Polar": "Yellow",
"Tmavá Hnědá Polarizovaná": "Brown",
"Fialová zrcadlová stříbrná vnitřní": "Purple|Silver",
"Cognac Sfumato Ninfea": "Brown",
"Světlá Hnědá Foto": "Brown",
"Hnědé Zrcadlová Uvnitř Šedá": "Brown|Grey",
"Grey Polarized Mirror Gold": "Grey|Gold",
"Tmavá Šedá Flash Stříbrná": "Grey|Silver",
"hnědá oranžová, 24 karátů, Iridium": "Multicolor",
"Hnědá zrcadlová vícevrstvá zlatá": "Brown|Gold",
"Light Green Silver Gradient": "Green|Silver",
"Černá čokoláda": "Brown",
"Light Brown Gradient Bordeaux": "Brown|Burgundy",
"Hnědá Foto": "Brown",
"Tmavě růžová polarizační": "Pink",
"Černá Polar": "Black",
"Avio Chiaro": "Blue",
"Lago": "Blue",
"Bruciato Polar": "Brown",
"Světlá Žlutá Gradient Hnědá": "Yellow|Brown",
"Tmavá Medová": "Brown",
"Limone/stříbrná": "Yellow|Silver",
"Purpurová/stříbrná": "Purple|Silver",
"Lente Camomilla Bio": "Yellow",
"Caffe Sfumato Specchio Oro": "Brown|Gold",
"Photogrey Extra Yellow": "Grey",
"zelená/hnědá": "Green|Brown",
"Hnědá gradientní bordó": "Brown|Burgundy",
"Ultravioletta Bio": "Purple",
"Champagne Flash Bio": "Yellow",
"modrá/šedá": "Blue|Grey",
"Polar Světlá Hnědá": "Brown",
"Ultravioletto": "Purple",
"Ardesia/Antiriflesso": "Grey",
"Graffite/Antiriflesso": "Grey",
"Mora": "Purple",
"Modrá Zrcadlová Stříbrná": "Blue|Silver",
"Tmavá Šedá Vnější strana Ar": "Grey",
"Tmavě modrá zrcadlově fialová": "Blue|Purple",
"Hnědá Tuning": "Brown",
"tmavě šedá, zrcadlové, modrá/červená": "Multicolor",
"Blue Multilayer Tuning": "Blue",
"Dark Grey Mirror Red Tuning": "Grey|Red",
"Tuning Dark Green": "Green",
"Modré ladění": "Blue",
"Light Green Tuning": "Green",
"Dark Blue Violet Mirror": "Blue|Purple",
"gradální, šedá, zrcadlové, stříbrná": "Grey|Silver",
"Břidlicová": "Grey",
"Tónování tmavě šedé zrcadlově červené": "Grey|Red",
"Zrcadlová Stříbrná": "Silver",
"Dark Grey Mirror Orange Tuning": "Grey|Orange",
"Tmavě hnědá zrcadlová": "Brown",
"Purpurová Zrcadlová Stříbrná": "Purple|Silver",
"Fialová/zelená": "Purple|Green",
"Šedá Gradient Tmavá Modrá": "Grey|Blue",
"Fialová Gradient Stříbrná": "Purple|Silver",
"Světlá Modrá Zrcadlová Stříbrná Šedá": "Multicolor",
"Červená/Modrá": "Red|Blue",
"Jednobarevná Černá": "Black",
"modrá/stříbrná, zrcadlové": "Blue|Silver",
"Red Polarized": "Red",
"Šedá/Stříbrná": "Grey|Silver",
"Gradient Purpurová": "Purple",
"Přechodová světle hnědá fialová": "Brown|Purple",
"Gradient Světlá Šedá": "Grey",
"Růžový gradient, odstín Lillac": "Pink|Purple",
"hnědá, gradální, zrcadlové": "Brown",
"Polarized Gradient Grey": "Grey",
"Polar Light Brown Gradient": "Brown",
"Gradient Brown Polarized": "Brown",
"Fialová Polar": "Purple",
"Gradient Polarized Blue": "Blue",
"Hnědá Růžová": "Brown|Pink",
"Gradient Světlá Fialová": "Purple",
"Modrá Šedá": "Blue|Grey",
"Purple Brown Mirror": "Purple|Brown",
"Gradientní odstín Lillac": "Purple",
"Šedá Gradient Fialová": "Grey|Purple",
"fialová, gradální": "Purple",
"Ultra černá": "Black",
"G-15 Zelená": "Green",
"Čirá/modrá": "Transparent|Blue",
"Čirá a modrá": "Transparent|Blue",
"Clear & Brown": "Transparent|Brown",
"zlato": "Gold",
"Světlá Hnědá Gradient Černá": "Brown|Black",
"B-15 hnědá": "Brown",
"stříbrná, s leskem": "Silver",
"polarizační, zelená Classic G-15": "Green",
"modrá, s leskem": "Blue",
"Čirá Šedá": "Transparent|Grey",
"světle modrá, gradální": "Blue",
"Fialová/Zlatá": "Purple|Gold",
"Světle šedá / tmavě modrá": "Grey|Blue",
"Hnědá/Modrá Gradient": "Brown|Blue",
"Šedá/Modrá Gradient": "Grey|Blue",
"modrá/černá": "Blue|Black",
"Hnědá vintage gradientní černá": "Brown|Black",
"Modrá vintage gradientní černá": "Blue|Black",
"Růžová/Černá": "Pink|Black",
"Růžová gradientní černá": "Pink|Black",
"Stříbrná/Modrá": "Silver|Blue",
"Šedá a modrá": "Grey|Blue",
"Šedá zrcadlová modrá/červená": "Multicolor",
"Zrcadlová šedá polarizovaná gradientní stříbrná": "Grey|Silver",
"Šedá Gradient Modrá": "Grey|Blue",
"Gradient Yellow/Blue": "Yellow|Blue",
"Modrá Hnědá": "Blue|Brown",
"Čirá/hnědá": "Transparent|Brown",
"Modrá Gradient Polar": "Blue",
"Polarizovaná růžová": "Pink",
"Stříbrná/Zelená": "Silver|Green",
"měděná, s leskem": "Orange",
"zelená, gradální, s leskem": "Green",
"zelená Classic G-15": "Green",
"stříbrná/růžová": "Silver|Pink",
"zelená s leskem": "Green",
"Cyclamen, s leskem": "Pink",
"Stříbrná/Hnědá": "Silver|Brown",
"Čirá Zrcadlová Bílá Zlatá": "Transparent|Gold",
"Tmavá Hnědá Zrcadlová Oranžová Zlatá": "Multicolor",
"Zelená Starožitná": "Green",
"růžová/hnědá, gradální": "Pink|Brown",
"měděná, zrcadlové": "Orange",
"Evolve Foto Modrá Do Fialové": "Blue|Purple",
"Evolve Foto Hnědá Do Tmavé Hnědé": "Brown",
"fotochromatické, zelená, gradální, modrá": "Green|Blue",
"Fialová zlatá": "Purple|Gold",
"Světlá Modrá Flash": "Blue",
"Zlatá/Červená": "Gold|Red",
"Hnědá A Šedá": "Brown|Grey",
"Hnědá a červená": "Brown|Red",
"zelená, zrcadlové, modrá": "Green|Blue",
"zelená Classic": "Green",
"Gradient Zlatá": "Gold",
"modrá/hnědá, gradální": "Blue|Brown",
"Fialová Zrcadlová Gradient Stříbrná": "Purple|Silver",
"Modrá zrcadlová gradientní stříbrná": "Blue|Silver",
"Světle zelená/stříbrná": "Green|Silver",
"Purpurová/zlatá": "Purple|Gold",
"Modrá/Zlatá": "Blue|Gold",
"Hnědá/Šedá": "Brown|Grey",
"Čirá/Šedá": "Transparent|Grey",
"Čirá a žlutá": "Transparent|Yellow",
"Světle modrá a fialová": "Blue|Purple",
"Čirá/Zelená": "Transparent|Green",
"Čirá/bílá": "Transparent",
"Oranžová Zrcadlová Růžová Polar": "Orange|Pink",
"Fialová a stříbrná": "Purple|Silver",
"Polar Šedá Zrcadlová Modrá": "Grey|Blue",
"polarizační, světle modrá": "Blue",
"Čirá/safírová": "Transparent|Blue",
"Žlutá Zrcadlová Flash Zlatá": "Yellow|Gold",
"Zrcadlově šedá gradientní polarizovaná šedá": "Grey",
"Hnědá zrcadlová červená / žlutá": "Multicolor",
"Modrá Gradient Tmavá Šedá": "Blue|Grey",
"safírová": "Blue",
"smaragd": "Green",
"ametyst": "Purple",
"Červená Hiper": "Red",
"Šedá zrcadlově červená": "Grey|Red",
"Polarizovaná Vínová": "Burgundy",
"Polarizovaná šedá zrcadlová modrá": "Grey|Blue",
"Hnědá Zrcadlová Zlatá Gradient": "Brown|Gold",
"hnědá/šedá, gradální": "Brown|Grey",
"červená, zrcadlové": "Red",
"Zelené zrcadlové stříbrné polarizované": "Green|Silver",
"hnědá/fialová": "Brown|Purple",
"tmavě fialová, Classic": "Purple",
"tmavě šedá, Classic": "Grey",
"Modrá Zrcadlová Zlatá Gradient": "Blue|Gold",
"Blue Chromance": "Blue",
"Šedá Chromance": "Grey",
"modrá, Classic": "Blue",
"Foto Šedá Zrcadlová Šedá": "Grey",
"zelená/modrá": "Green|Blue",
"hnědá/zlatá": "Brown|Gold",
"Tmavě fialová/červená": "Purple|Red",
"Stříbrná/Šedá": "Silver|Grey",
"Světlá fialová / růžovozlatá": "Multicolor",
"okr": "Brown",
"Šedá Zrcadlová Polar": "Grey",
"Zelená Zrcadlová Modrá Polar": "Green|Blue",
"Zelená Chromance": "Green",
"Fialová Chromance": "Purple",
"Hnědá Chromance": "Brown",
"Zelená Gradient Hnědá": "Green|Brown",
"Tmavá Hnědá - Polar": "Brown",
"Polar Bronzová": "Orange",
"Polar Červená": "Red",
"Polar Modrá": "Blue",
"Polarizovaná Tmavá Hnědá Classic": "Brown",
"Polarizovaná Tmavá Zelená Classic": "Green",
"Polarizovaná Tmavá Šedá Classic": "Grey",
"Polarizovaná křišťálově modrá": "Blue",
"Hnědá Zrcadlová Tmavě Červená": "Brown|Red",
"polarizační, tmavě hnědá": "Brown",
"Polar Růžová": "Pink",
"světle zlatá": "Gold",
"měděná": "Orange",
"Světle Šedá/Tmavě Šedá": "Grey",
"Hnědá/Oranžová": "Brown|Orange",
"Tmavá zlatá": "Gold",
"Modrá a fialová": "Blue|Purple",
"šedá/fialová": "Grey|Purple",
"Modrá přecházející do šedé a zrcadlově růžové": "Multicolor",
"Čirá Gradient Červená Zrcadlová Červená": "Transparent|Red",
"Fialová Gradient Hnědá": "Purple|Brown",
"Růžová Šedá": "Pink|Grey",
"Lila světle šedá": "Purple|Grey",
"Šedá Gradient Zrcadlová Červená": "Grey|Red",
"Hnědá/Tmavě hnědá": "Brown",
"Fialová/Šedá": "Purple|Grey",
"Světle šedá hnědá": "Grey|Brown",
"Čirá Gradient Modrá Zrcadlová Červená": "Multicolor",
"světle modrá, zrcadlové": "Blue",
"Hnědá Zrcadlová Stříbrná": "Brown|Silver",
"zelená, zrcadlové": "Green",
"Red Hiper 8 Mirror": "Red",
"Šeříková Flash": "Purple",
"Gradient Tmavá Fialová": "Purple",
"Čirá a tmavě fialová": "Transparent|Purple",
"Light Blue Lenses": "Blue",
"Grey/Silver Mirror": "Grey|Silver",
"šedorůžová, gradální": "Grey|Pink",
"Přechodová hnědo-šedá zrcadlová zlatá": "Multicolor",
"Gradient Hnědá Zrcadlová Zlatá": "Brown|Gold",
"Světlá Šedá Zrcadlová Stříbrná": "Grey|Silver",
"Jednobarevná Purpurová": "Purple",
"Gradient Hnědá/Růžová": "Brown|Pink",
"Šedá Gradient Oranžová": "Grey|Orange",
"Oranžová Gradient Šedá": "Orange|Grey",
"Světlá Gradient Hnědá": "Brown",
"Gradient Dark Purple/Grey": "Purple|Grey",
"Hnědá Gradient Zrcadlová Zlatá": "Brown|Gold",
"Šedá Gradient Flash": "Grey",
"Trigradient Purple/Blue/Pink": "Multicolor",
"Oystershell Rose/Mauve": "Pink|Purple",
"Jednobarevná Zelená": "Green",
"Trojitý gradient, světle hnědá/fialová": "Brown|Purple",
"Dark Grey Mirror Green": "Grey|Green",
"stříbrná gradální, s leskem": "Silver",
"Transition Green": "Green",
"Růžová vnitřní zrcadlová stříbrná": "Pink|Silver",
"Azurová Uvnitř Zrcadlová Stříbrná": "Blue|Silver",
"Transition Light Grey To Dark Grey": "Grey",
"žlutá, gradální": "Yellow",
"Foto Zelená": "Green",
"Světle žlutá zrcadlová stříbrná vnitřní": "Yellow|Silver",
"Světle zelená záblesková stříbrná": "Green|Silver",
"Světlá Modrá Zrcadlová Zelená": "Blue|Green",
"Modrá Zrcadlová Purpurová Růže": "Multicolor",
"Hnědá Fotochromatická": "Brown",
"Transition Pink": "Pink",
"Gradientní hnědá, gradientní růžová": "Brown|Pink",
"Fotochromatická azurová": "Blue",
"Růžová přechodová": "Pink",
"Růžová/fialová": "Pink|Purple",
"Světle zelená zrcadlová stříbrná vnitřní": "Green|Silver",
"Žlutá Zrcadlová": "Yellow",
"Čiré zrcadlo – pravá platina": "Transparent|Silver",
"Gradientní hnědá/modrá/fialová": "Multicolor",
"Gradientní světle hnědá/modrá/fialová": "Multicolor",
"Grey Tampo Swarovski Gold": "Grey|Gold",
"Brown Swarovski Gold": "Brown|Gold",
"Purple Brown Tampo Swarovski Pink": "Multicolor",
"Foto Zelená až Tmavá Zelená": "Green",
"Světlá Hnědá Gradient Modrá": "Brown|Blue",
"Oranžová Gradient Fialová": "Orange|Purple",
"Azurová Gradient Tmavá Modrá Polar": "Blue",
"Gradient Růžová Zrcadlová Oranžová": "Pink|Orange",
"Světlá Fialová Gradient Šedá": "Purple|Grey",
"Dark Pink Mirror Gold": "Pink|Gold",
"Modrá stříbřitá, zrcadlový odstín, vnitřek": "Blue",
"Azure/Dark Grey": "Blue|Grey",
"Šedá Zrcadlová Mléčná Modrá": "Grey|Blue",
"Světle azurová zrcadlová stříbrná": "Blue|Silver",
"Světle fialová zrcadlová vnitřní": "Purple|Silver",
"Gradient Green Bordeaux": "Green|Burgundy",
"Zelená Modrá Gradient": "Green|Blue",
"modrá Tiffany, gradální": "Blue",
"Světle Šedá Gradient Zelená": "Grey|Green",
"Fialová Zrcadlová Stříbrná": "Purple|Silver",
"Fialová Gradient Šedá": "Purple|Grey",
"Růžová zrcadlová modrá": "Pink|Blue",
"Světle hnědo-modrá": "Brown|Blue",
"Pink Back Mirror Silver": "Pink|Silver",
"Šedá Gradient": "Grey",
"Azure/Burgundy": "Blue|Burgundy",
"Azurová Gradient Hnědá": "Blue|Brown",
"Světle oranžová zrcadlová stříbrná vnitřní": "Orange|Silver",
"Azure/Violet": "Blue|Purple",
"Azure Gradient Blue Polarized": "Blue",
"Světle Růžová Zrcadlová Gradient Stříbrná": "Pink|Silver",
"Tmavá Šedá Zrcadlová Stříbrná Polarizovaná": "Grey|Silver",
"Světle Šedá Gradient Šedá": "Grey",
"tmavě šedá, zrcadlové, zlatá": "Grey|Gold",
"polarizační, šedá, zrcadlové, stříbrná": "Grey|Silver",
"Pink Gradient Grey Flash Silver": "Multicolor",
"Růžová Gradient Modrá": "Pink|Blue",
"Fotochromatická růžová": "Pink",
"Dark Grey/Mirror Gold": "Grey|Gold",
"Bronze Mirror Gold": "Orange|Gold",
"Pink Mirror gradientní růžová": "Pink",
"Světlá Hnědá Zrcadlová Stříbrná": "Brown|Silver",
"Světlá Hnědá Zrcadlová Stříbrná Gradient": "Brown|Silver",
"Čirá polooranžová/tmavě růžová": "Multicolor",
"Čirá pologradientní oranžová fialová": "Multicolor",
"Brown 24K Iridium": "Brown|Gold",
"Světlá Hnědá Tmavá Hnědá Gradient": "Brown",
"žlutá, gradální, fialová": "Yellow|Purple",
"Růžová Zrcadlová Zlatá": "Pink|Gold",
"Světlá/Tmavá Hnědá Gradient": "Brown",
"Světlá Žlutá Gradient Okrová": "Yellow|Brown",
"Yellow Mirror Internal Silver": "Yellow|Silver",
"Růžová Gradient Růžová": "Pink",
"Grey Mirror Black/Mirror Blue Violet": "Multicolor",
"Dark Grey/Mirror Silver": "Grey|Silver",
"Hnědá Oranžová Metalická": "Brown|Orange",
"Šedá Zrcadlová Gradient Stříbrná": "Grey|Silver",
"světle hnědá, zrcadlové, stříbrná, gradální": "Brown|Silver",
"Tmavě Šedá Gradient Polar": "Grey",
"Fialová Tmavá Šedá": "Purple|Grey",
"Gradient Světlá Modrá Zrcadlová Stříbrná": "Blue|Silver",
"Azurová Gradient Růžová Gradient Hnědá": "Multicolor",
"Čirá a Ar": "Transparent",
"Azure Grad Pink Grad Brown": "Multicolor",
"Trojpřechodová hnědá / fialová / modrá": "Multicolor",
"Přechodová zeleno-tmavě hnědá": "Green|Brown",
"Brown Gradient Purple Black": "Multicolor",
"Modrá Polar Flash Stříbrná": "Blue|Silver",
"Purpurová Polar": "Purple",
"Polarizační přechodová šedo-hnědá": "Grey|Brown",
"Polarizační růžová zrcadlová stříbrná": "Pink|Silver",
"Šedá Gradient Hnědá Polar": "Grey|Brown",
"Polarizační přechodová modro-hnědá": "Blue|Brown",
"Polar Šedá Gradient Fialová": "Grey|Purple",
"Brown/Violet/Blue": "Multicolor",
"Polar Šedá Gradient Tmavá Fialová": "Grey|Purple",
"Šedá Gradient Tmavá Šedá Polar": "Grey",
"Polarized Grey Gradient Brown": "Grey|Brown",
"Polarizační přechodová šedo-fialová": "Grey|Purple",
"Šedá Gradient Tmavá Šedá": "Grey",
"Hnědá Gradient Purpurová Gradient Černá": "Multicolor",
"Gradient Hnědá/Fialová/Modrá": "Multicolor",
"Gradient Dark Blue": "Blue",
"Gradient Tmavá Šedá": "Grey",
"Růžová Tmavá Zrcadlová Červená": "Pink|Red",
"Bronzová Polar": "Orange",
"Grey Gradient Blue Polar": "Grey|Blue",
"Tmavě Fialová/Šedá": "Purple|Grey",
"Šedá levandulová": "Grey",
"zrcadlově růžovo-stříbrná": "Pink|Silver",
"Modrá Stříbrná Zrcadlová": "Blue|Silver",
"Světlá Hnědá Gradient Světlá Zelená": "Brown|Green",
"Šedá Zrcadlová Žlutá Růžová": "Multicolor",
        #marcolin
        "NONE":"",
        "BLACK": "Black",
        "RED": "Red",
        "BLUE": "Blue",
        "LIGHT BLUE": "Blue",
        "GREEN": "Green",
        "GOLD": "Gold",
        "SILVER": "Silver",
        "WHITE": "Grey",
        "COPPER": "Orange",
        "BRONZE": "Orange",
        "ORANGE": "Orange",
        "YELLOW": "Yellow",
        "VIOLET": "Purple",
        "BROWN": "Brown",
        "PINK": "Pink",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "SMOKE": "Transparent",
        "BURGUNDY": "Burgundy",
        #kering
        "NONE":"",
        "BLACK": "Black",
        "RED": "Red",
        "BLUE": "Blue",
        "LIGHT BLUE": "Blue",
        "GREEN": "Green",
        "GOLD": "Gold",
        "SILVER": "Silver",
        "WHITE": "Grey",
        "COPPER": "Orange",
        "BRONZE": "Orange",
        "ORANGE": "Orange",
        "YELLOW": "Yellow",
        "VIOLET": "Purple",
        "BROWN": "Brown",
        "PINK": "Pink",
        "MULTICOLOR": "Multicolor",
        "TRANSPARENT": "Transparent",
        "SMOKE": "Transparent",
        "BURGUNDY": "Burgundy",
    },
    "Glasses_lens_material": {
        #safilo
        "ONLY FOR USA - NOT DEFINED":"",
        "0":"",
        "BIO NYLON LENS": "Nylon",
        "NYLON LENS": "Nylon",
        "POLICARBONATE LENS": "Polycarbonate",
        "CR39": "CR 39",
        "GLASS LENS": "Glass",
        "PMMA": "Plastic",
        "POLYESTER": "Plastic",
        "ECO CO-POLYESTER": "Plastic",
        "TRIACETATE LENS": "Plastic",
        #luxottica
        "Cr39": "CR 39",
        "MINERÁLNÍ": "Glass",
        "sklo": "Glass",
        "PLAST": "Plastic",
        "bio polyamid": "Plastic",
        "polyamid": "Plastic",
        "jiné než sklo": "Plastic",
        "Nxt": "Plastic",
        "Acrylic": "Plastic",
        #marcolin
        "ACETATE":"Plastic",
        "NYLON":"Nylon",
        "BIO NYLON": "Nylon",
        "POLAR BIO NYLON": "Nylon",
        "POLAR NYLON": "Nylon",
        "POLAR CR 39": "Polar CR 39",
        "POLYCARBONATE": "Polycarbonate",
        "CR 39": "CR 39",
        "POLAR PC": "Polar PC",
        "GLASS": "Glass",
        "POLAR GLASS": "Glass",
        "NXT LENS": "Plastic",
        "RECYCLED PMMA DEMO LENS": "Plastic",
        "STG": "Plastic",
        "BRILLIANT": "Plastic",
        "PURE": "Plastic",
        "MAUI ULTRA": "Plastic",
        "PURE LT": "Plastic",
        "ELLUME": "Plastic",
        #kering
        "ACETATE":"Plastic",
        "NYLON":"Nylon",
        "BIO NYLON": "Nylon",
        "POLAR BIO NYLON": "Nylon",
        "POLAR NYLON": "Nylon",
        "POLAR CR 39": "Polar CR 39",
        "POLYCARBONATE": "Polycarbonate",
        "CR 39": "CR 39",
        "POLAR PC": "Polar PC",
        "GLASS": "Glass",
        "POLAR GLASS": "Glass",
        "NXT LENS": "Plastic",
        "RECYCLED PMMA DEMO LENS": "Plastic",
        "STG": "Plastic",
        "BRILLIANT": "Plastic",
        "PURE": "Plastic",
        "MAUI ULTRA": "Plastic",
        "PURE LT": "Plastic",
        "ELLUME": "Plastic",
    },
    "Glasses_lens_effect": {
        "": "",
        #safilo
        "BLUE FILTER + ANTIREFLEX": "",
        "BLUE FILTER": "",
        "CILIARY NET": "",
        "DECOR": "",
        "HIGH CONTRAST ANTIREFLEX": "",
        "HIGH CONTRAST BLUE FILTER": "",
        "HIGH CONTRAST MULTILAYER": "",
        "HIGH CONTRAST": "",
        "INFRARED": "",
        "MIRROR + OLEOPHOBIC": "Mirror",
        "MULTILAYER + OLEOPHOBIC": "",
        "MULTILAYER": "",
        "OLEOPHOBIC": "",
        "SLEWING": "",
        "SOLID TINT": "",
        "ANTIREFLEX":"",
        "0": "",
        "GRADIENT + MIRROR": "Gradient!Mirror",
        "MIRROR": "Mirror",
        "DOUBLE GRADIENT": "Gradient",
        "FLASH MIRROR": "Mirror",
        "TRIPLE GRADIENT": "Gradient",
        "GRADIENT + ANTIREFLEX": "Gradient",
        "GRADIENT + FULL MIRROR": "Mirror",
        "GRADIENT MIRROR": "Mirror",
        "HIGH CONTRAST MIRROR ANTIREFLEX": "Mirror",
        "MIRROR + ANTIREFLEX": "Mirror",
        "HIGHT CONTRAST MIRROR": "Mirror",
        "MIRROR DECOR": "Mirror",
        #luxottica
        "gradient":"Gradient",
        "mirror":"Mirror",
        "zrcadlo":"Mirror",
        "přech":"Gradient",
        #marcolin
        "ANTIFOG": "",
        "AR (ANTI RIFLESSO)": "",
        "BLUE NEON": "",
        "MULTI TREATMENT": "",
        "NONE": "",
        "DEMO LENS": "",
        "ACETATE": "Plastic",
        "LEATHER": "",
        "RECYCLED FABRIC": "",
        "": "",
        "": "",
        "": "",
        "": "",

        "MIRROR": "Mirror",
        "MIRROR(DOUBLE)": "Mirror",
        "FLASH": "Mirror",
        "AR+MIRROR": "Mirror",
        "GRADIENT+MIRROR": "Gradient|Mirror",
        "GRADIENT FLASH": "Gradient",
        "GRADIENT(DOUBLE)": "Gradient",
        "PHOTOCROMATIC": "Photochromic",
        #kering
        "MIRROR": "Mirror",
        "MIRROR(DOUBLE)": "Mirror",
        "FLASH": "Mirror",
        "AR+MIRROR": "Mirror",
        "GRADIENT+MIRROR": "Gradient|Mirror",
        "GRADIENT FLASH": "Gradient",
        "GRADIENT(DOUBLE)": "Gradient",
        "PHOTOCROMATIC": "Photochromic",
    },
    "Sunglasses_filter": {
        "": "",
        #safilo
        "0":"",
        #luxottica
        #marcolin
        "0": "Category 0",
        "1": "Category 1",
        "2": "Category 2",
        "3": "Category 3",
        "3-3": "Category 3",
        "4": "Category 4",
        "0-1": "Category range 0 - 1",
        "0-2": "Category range 0 - 2",
        "0-3": "Category range 0 - 3",
        "0-4": "Category range 0 - 4",
        "1-2": "Category range 1 - 2",
        "1-3": "Category range 1 - 3",
        "1-4": "Category range 1 - 4",
        "2-3": "Category range 2 - 3",
        "2-4": "Category range 2 - 4",
        "3-4": "Category range 3 - 4",
        #kering
        "0": "Category 0",
        "1": "Category 1",
        "2": "Category 2",
        "3": "Category 3",
        "4": "Category 4",
        "0-1": "Category range 0 - 1",
        "0-2": "Category range 0 - 2",
        "0-3": "Category range 0 - 3",
        "0-4": "Category range 0 - 4",
        "1-2": "Category range 1 - 2",
        "1-3": "Category range 1 - 3",
        "1-4": "Category range 1 - 4",
        "2-3": "Category range 2 - 3",
        "2-4": "Category range 2 - 4",
        "3-4": "Category range 3 - 4",
    },
    "Glasses_gendre": {
        "": "",
        #safilo
        "UNISEX ADULT":"Man|Woman",
        "MAN": "Man",
        "WOMAN": "Woman",
        "YOUNG KIDS (4-6)": "Child",
        "GIRLTEEN (11-15)": "Child",
        "INFANT AND TODDLERS (0-3)": "Child",
        "BOYTEEN (11-15)": "Child",
        "UNISEX TEENAGER (11-15)": "Child",
        "JUNIOR (7-10)": "Child",
        #luxottica
        "Muž": "Man",
        "Žena": "Woman",
        "Unisex": "Man|Woman",
        #marcolin
        "WOMAN": "Woman",
        "MAN": "Man",
        "UNISEX": "Man|Woman",
        "KID": "Child",
        #kering
        "WOMAN": "Woman",
        "MAN": "Man",
        "UNISEX": "Man|Woman",
        "KID": "Child",
    },
}
# ==========================================
# 🗺️ THE CONFIGURATION (Global -> Manufacturer)
# ==========================================
MANUFACTURER_CONFIG = {
    "safilo": {
        "file": "safilo.xlsx",
        "brands": ["Carrera", "Polaroid", "Smith", "Boss", "Tommy Hilfiger", "Fossil", "Pierre Cardin", "Marc Jacobs"],
        "columns": {
            "Combination": "Size",
            "Barcode": "EAN/UPC",
            "Glasses_type": "Product Type Desc.",
            "Manufacturer": "Brand Description",
            "Glasses_size_temple_length": "Temple Length",
            "Glasses_size_lens_height": "Lens Height",
            "Glasses_size_lens_width": "Size",
            "Glasses_size_bridge": "Bridge Length",
            "Glasses_shape": "Shape",
            "Glasses_other_info": "Shape",
            "Glasses_frame_type": "Description Rym type",
            "Frame_Colour": "Description Color Family 1",
            "Temple_Colour": "Description Color Family 1",
            "Glasses_main_material": "Group material",
            "Glasses_lens_Colour": "Description Lens Color Family",
            "Glasses_lens_material": "Lens Material Description",
            "Glasses_lens_effect": ["Polarized", "Photochromic", "Treatement Description"],
            "Polarized_raw": "Polarized",
            "Photochromic_raw": "Photochromic",
            "Treatement_Description_raw": "Treatement Description",
            "Sunglasses_filter": "Transparency (%)",
            "Glasses_gendre": "Gender",
            "Glasses_model": "Model",
            "Glasses_color_code": "COLOUR CODE",
            "SunGlasses_RX_lenses": "RX able INT",
            "Brand": "Brand Description",
            "Producing_company": ""      
        }
    },
    "luxottica": {
        "file": "luxottica.xlsx",
        "brands": ["Ray-Ban", "Oakley", "Persol", "Prada", "Versace", "Burberry", "Dolce & Gabbana", "Michael Kors", "Vogue", "Arnette", "Ralph"],
        "columns": {
            "Combination": "Velikost",
            "Barcode": "UPC",
            "Glasses_type": "Kolekce",
            "Manufacturer": "Název značky",
            "Glasses_size_temple_length": "Délka stranice",
            "Glasses_size_lens_height": "Výška čočky",
            "Glasses_size_lens_width": "Velikost",
            "Glasses_size_bridge": "Velikost nosníku",
            "Glasses_shape": "Tvar",
            "Glasses_other_info": "Flex",
            "Glasses_frame_type": "Typ",
            "Frame_Colour": "Barva očnice.1", 
            "Temple_Colour": "Barva očnice.1",
            "Glasses_main_material": "Materiál očnice",
            "Glasses_lens_Colour": "Barva čočky",
            "Glasses_lens_material": "Materiál čočky",
            "Glasses_lens_effect": ["Polarizované", "Fotochromatické", "Barva čočky"],
            "Polarizovane_raw": "Polarizované",
            "Fotochromaticke_raw": "Fotochromatické",
            "Barva_cocky_raw": "Barva čočky",
            "Glasses_gendre": "Pohlaví",
            "Glasses_usable": "Motiv",
            "Glasses_collection": "Skládací",
            "Glasses_model": "Kód modelu",
            "Glasses_color_code": "Kód barvy",
            "Brand": "Název značky",
            "Producing_company": "" 
        }
    },
    "kering": {
        "file": "kering.xlsx",
        "brands": ["Gucci", "Saint Laurent", "Balenciaga", "Montblanc", "Bottega Veneta", "Alexander McQueen", "Dunhill", "Puma"],
        "columns": {
            "Combination": "Size 1",
            "Barcode": "UPC code(Z2)",
            "Glasses_type": "Material Group",
            "Manufacturer": "Brand", 
            "Glasses_size_temple_length": "Temple length (mm)",
            "Glasses_size_lens_height": "Lens Height",
            "Glasses_size_lens_width": "Size 1",
            "Glasses_size_bridge": "Bridge Length (mm)",
            "Glasses_shape": "Shape Description",
            "Glasses_other_info": ["Hinge description", "Flex"],
            "Glasses_frame_type": "Rim",
            "Frame_Colour": "Front Main Color Description",
            "Temple_Colour": "Temple Main Color Description",
            "Glasses_main_material": "Front Main Material Description",
            "Glasses_lens_Colour": "Lens Main Color Description",
            "Glasses_lens_material": "Lens Material Description",
            "Glasses_lens_effect": ["Polarized Lens", "Photocromic", "Lens Effect Description"],
            "Polarized_Lens_raw": "Polarized Lens",
            "Photocromic_raw": "Photocromic",
            "Lens_Effect_Description_raw": "Lens Effect Description",
            "Sunglasses_filter": "Filter Category",
            "Glasses_gendre": "Fashion Attribute 2",
            "Glasses_collection": "Foldable",
            "SunGlasses_RX_lenses": "Convertible in optical frame",
            "Brand": "Brand",
            "Case_weight_g": "Gross Weight",
            "Glasses_weight_g": "Net Weight",
            "Item_origin_country": "Country of origin",
            "Producing_company": "",
            "Family_descriptions_raw": "Family descriptions",
        }
    },
    "marcolin": {
        "file": "marcolin.xlsx",
        "brands": ["Tom Ford", "Guess", "Adidas", "Max Mara", "Moncler", "Zegna", "Gant", "Harley Davidson", "Skechers", "Web", "Timberland"],
        "columns": {
            "Combination": "Size 1",
            "Barcode": "UPC code(Z2)",
            "Glasses_type": "Material Group",
            "Manufacturer": "Brand", 
            "Glasses_size_temple_length": "Temple length (mm)",
            "Glasses_size_lens_height": "Lens Height",
            "Glasses_size_lens_width": "Size 1",
            "Glasses_size_bridge": "Bridge Length (mm)",
            "Glasses_shape": "Shape Description",
            "Glasses_other_info": ["Hinge description", "Flex"],
            "Glasses_frame_type": "Rim",
            "Frame_Colour": "Front Main Color Description",
            "Temple_Colour": "Temple Main Color Description",
            "Glasses_main_material": "Front Main Material Description",
            "Glasses_lens_Colour": "Lens Main Color Description",
            "Glasses_lens_material": "Lens Material Description",
            "Glasses_lens_effect": ["Polarized Lens", "Photocromic", "Lens Effect Description"],
            "Polarized_Lens_raw": "Polarized Lens",
            "Photocromic_raw": "Photocromic",
            "Lens_Effect_Description_raw": "Lens Effect Description",
            "Sunglasses_filter": "Filter Category",
            "Glasses_gendre": "Fashion Attribute 2",
            "Glasses_collection": "Foldable",
            "SunGlasses_RX_lenses": "Convertible in optical frame",
            "Brand": "Brand",
            "Case_weight_g": "Gross Weight",
            "Glasses_weight_g": "Net Weight",
            "Item_origin_country": "Country of origin",
            "Producing_company": "",
            "Family_descriptions_raw": "Family descriptions",
        }
    }
}

# ==========================================
# 📥 THE LOADER
# ==========================================
@st.cache_data(show_spinner=False)
def load_all_catalogs(config):
    virtual_catalog = {}
    current_dir = os.getcwd()
    
    for mfg_name, settings in config.items():
        file_name = settings["file"]
        file_path = os.path.join(current_dir, file_name)
        
        if not os.path.exists(file_path):
            st.warning(f"⚠️ Missing File: '{file_name}'")
            continue
            
        try:
            if file_name.endswith('.csv'):
                try:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=',')
                except:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=';')
            else:
                df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
                
            df.columns = df.columns.astype(str).str.strip()
            
            new_cols = []
            seen = {}
            for c in df.columns:
                if c in seen:
                    seen[c] += 1
                    new_cols.append(f"{c}.{seen[c]}")
                else:
                    seen[c] = 0
                    new_cols.append(c)
            df.columns = new_cols
            
        except Exception as e:
            st.error(f"❌ Error loading {file_name}: {e}")
            continue

        new_df = pd.DataFrame()
        
        for global_name, mfg_names in settings["columns"].items():
            if not mfg_names: continue
            if isinstance(mfg_names, str):
                mfg_names = [mfg_names]
                
            existing_cols = [col for col in mfg_names if col in df.columns]
            
            if existing_cols:
                if len(existing_cols) == 1:
                    col_data = df[existing_cols[0]]
                    if isinstance(col_data, pd.DataFrame):
                        col_data = col_data.iloc[:, 0]
                    new_df[global_name] = col_data
                else:
                    def merge_row(row):
                        vals = [str(row[c]).strip() for c in existing_cols if pd.notna(row[c]) and str(row[c]).strip().lower() not in ("nan", "")]
                        return ", ".join(vals) if vals else ""
                    new_df[global_name] = df.apply(merge_row, axis=1)

        # ==========================================
        # 🧠 CUSTOM RULES ENGINE & STRICT TRANSLATOR
        # ==========================================
        
        # We will store unknown values here to report to the user later
        if 'unmapped_values' not in st.session_state:
            st.session_state.unmapped_values = set()

        def process_cell_strict(row, col_name, mfg):
            final_values = set()
            raw_val = str(row.get(col_name, "")).strip()
            
            # --- 1. CUSTOM RULES ENGINE ---
            if col_name == "Glasses_other_info":
                if mfg == "safilo":
                    if pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper():
                        final_values.add("Flex")
                
                elif mfg == "luxottica":
                    raw_info = str(row.get("Glasses_other_info", "")).strip().upper()
                    if raw_info == "X":
                        final_values.add("Flex")
                    if pd.notna(row.get("Glasses_collection")) and str(row["Glasses_collection"]).strip().upper() == "X":
                        final_values.add("Flexible glasses")
                
                elif mfg in ["kering", "marcolin"]:
                    if pd.notna(row.get("Family_descriptions_raw")):
                        if "double bridge" in str(row["Family_descriptions_raw"]).lower():
                            final_values.add("Double bridge")

            # 🔥 NEW: GLASSES LENS EFFECT ENGINE 🔥
            elif col_name == "Glasses_lens_effect":
                if mfg == "safilo":
                    if str(row.get("Polarized_raw", "")).strip().upper() == "X":
                        final_values.add("Polarized")
                    if str(row.get("Photochromic_raw", "")).strip().upper() == "X":
                        final_values.add("Photochromic")
                    raw_eff = str(row.get("Treatement_Description_raw", "")).strip()
                    if raw_eff and raw_eff.lower() != "nan":
                        t_dict = VALUE_TRANSLATOR.get(col_name, {})
                        l_dict = {str(k).lower(): v for k, v in t_dict.items() if k}
                        for p in [x.strip() for x in raw_eff.split(",") if x.strip()]:
                            if p.lower() in l_dict:
                                if l_dict[p.lower()]: final_values.add(l_dict[p.lower()])
                            else:
                                st.session_state.unmapped_values.add(f"Safilo -> {col_name}: '{p}'")

                elif mfg == "luxottica":
                    if str(row.get("Polarizovane_raw", "")).strip().upper() == "X":
                        final_values.add("Polarized")
                    if str(row.get("Fotochromaticke_raw", "")).strip().upper() == "X":
                        final_values.add("Photochromic")
                    raw_eff = str(row.get("Barva_cocky_raw", "")).strip()
                    if raw_eff and raw_eff.lower() != "nan":
                        matched = False
                        t_dict = VALUE_TRANSLATOR.get(col_name, {})
                        for kw, m_val in t_dict.items():
                            if kw and kw.lower() in raw_eff.lower():
                                if m_val: final_values.add(m_val)
                                matched = True
                        if not matched:
                            st.session_state.unmapped_values.add(f"Luxottica -> {col_name} (Keyword Search): '{raw_eff}'")

                elif mfg in ["kering", "marcolin"]:
                    if str(row.get("Polarized_Lens_raw", "")).strip().upper() == "X":
                        final_values.add("Polarized")
                    if str(row.get("Photocromic_raw", "")).strip().upper() == "YES":
                        final_values.add("Photochromic")
                    raw_eff = str(row.get("Lens_Effect_Description_raw", "")).strip()
                    if raw_eff and raw_eff.lower() != "nan":
                        t_dict = VALUE_TRANSLATOR.get(col_name, {})
                        l_dict = {str(k).lower(): v for k, v in t_dict.items() if k}
                        for p in [x.strip() for x in raw_eff.split(",") if x.strip()]:
                            if p.lower() in l_dict:
                                if l_dict[p.lower()]: final_values.add(l_dict[p.lower()])
                            else:
                                st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name}: '{p}'")
                
                return ", ".join(sorted(list(final_values)))

            # --- 🕶️ RX LENSES ENGINE ---
            elif col_name == "SunGlasses_RX_lenses":
                raw_rx = str(row.get(col_name, "")).strip().upper()
                
                # Safilo, Kering, and Marcolin all use "X" for Yes
                if mfg in ["safilo", "kering", "marcolin"]:
                    if raw_rx == "X":
                        final_values.add("Yes")
                        
                # Luxottica rule pending (from To-Do list)
                elif mfg == "luxottica":
                    pass 
                
                return ", ".join(sorted(list(final_values)))
            
           # --- ☀️ SUNGLASSES FILTER ENGINE (Safilo Only) ---
            elif col_name == "Sunglasses_filter" and mfg == "safilo":
                raw_val = str(row.get(col_name, "")).strip()
                
                if raw_val and raw_val.lower() != "nan":
                    clean_numbers = re.findall(r'\d+\.?\d*', raw_val)
                    matched_by_math = False
                    
                    # 1. Try the Math Engine first
                    if clean_numbers:
                        vlt = float(clean_numbers[0])
                        if 80 <= vlt <= 100:
                            final_values.add("Category 0")
                            matched_by_math = True
                        elif 43 <= vlt < 80:
                            final_values.add("Category 1")
                            matched_by_math = True
                        elif 18 <= vlt < 43:
                            final_values.add("Category 2")
                            matched_by_math = True
                        elif 8 <= vlt < 18:
                            final_values.add("Category 3")
                            matched_by_math = True
                        elif 3 <= vlt < 8:
                            final_values.add("Category 4")
                            matched_by_math = True
                            
                    # 2. If math failed (no numbers or out of range), fallback to the Dictionary
                    if not matched_by_math:
                        if col_name in VALUE_TRANSLATOR:
                            translation_dict = VALUE_TRANSLATOR[col_name]
                            lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                            parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                            
                            for part in parts:
                                part_lower = part.lower()
                                if part_lower in lower_dict:
                                    if lower_dict[part_lower]: # Add if not banned ("")
                                        final_values.add(lower_dict[part_lower])
                                else:
                                    # Still unmapped? Now we flag it!
                                    st.session_state.unmapped_values.add(f"Safilo -> {col_name}: '{part}'")
                        else:
                            final_values.add(raw_val)
                            
                return ", ".join(sorted(list(final_values)))

            # --- 2. KEYWORD SUBSTRING MATCHER (Luxottica Lens Color) ---
            elif col_name == "Glasses_lens_Colour" and mfg == "luxottica":
                if raw_val and raw_val.lower() != "nan":
                    matched = False
                    if col_name in VALUE_TRANSLATOR:
                        translation_dict = VALUE_TRANSLATOR[col_name]
                        for keyword, mapped_val in translation_dict.items():
                            if keyword and keyword.lower() in raw_val.lower():
                                if mapped_val: final_values.add(mapped_val)
                                matched = True
                    if not matched:
                        st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name} (Keyword Search): '{raw_val}'")
                return ", ".join(sorted(list(final_values)))

            # --- 3. STRICT DICTIONARY TRANSLATOR (Everything else) ---
            elif raw_val and raw_val.lower() != "nan":
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                    parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                    
                    for part in parts:
                        part_lower = part.lower()
                        if part_lower in lower_dict:
                            if lower_dict[part_lower]: final_values.add(lower_dict[part_lower])
                        else:
                            st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
                else:
                    final_values.add(raw_val)

            return ", ".join(sorted(list(final_values)))

        # Apply the Engine
        for target_col in new_df.columns:
            if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses"]:
                new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

        # ZERO-STRIPPER
        if "Barcode" in new_df.columns:
            new_df["join_key"] = new_df["Barcode"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
            new_df = new_df[new_df["join_key"].notna() & (new_df["join_key"] != "nan") & (new_df["join_key"] != "")]
        else:
            st.error(f"❌ CRITICAL: 'Barcode' missing in {mfg_name} after extraction.")

        new_df["Producing_company"] = mfg_name.title()

        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = new_df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION & UI
# ==========================================

st.sidebar.header("Control Panel")

if st.sidebar.button("🗑️ Clear Memory & Reload Data", type="primary"):
    st.cache_data.clear()
    st.session_state.clear() 
    st.rerun()

with st.spinner("Building Virtual Catalog from scratch..."):
    catalog = load_all_catalogs(MANUFACTURER_CONFIG)

if not catalog:
    st.warning("No manufacturer catalogs loaded. Fix errors before proceeding.")
    st.stop()

@st.cache_data(show_spinner=False)
def get_master_database(cat):
    all_dfs = list(cat.values())
    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df.drop_duplicates(subset=['join_key'], keep='first', inplace=True)
    master_df.set_index('join_key', inplace=True)
    return master_df

master_db = get_master_database(catalog)

# 🚨 REPORT UNMAPPED VALUES
if 'unmapped_values' in st.session_state and st.session_state.unmapped_values:
    st.warning("⚠️ Action Required: Unmapped Values Found! The following values are not in your dictionary and were ignored.")
    
    # Group the errors by Manufacturer
    unmapped_grouped = {}
    for error in st.session_state.unmapped_values:
        if " -> " in error:
            mfg, detail = error.split(" -> ", 1)
        else:
            mfg, detail = "Other", error
            
        if mfg not in unmapped_grouped:
            unmapped_grouped[mfg] = []
        unmapped_grouped[mfg].append(detail)
        
    # Generate a separate rollout (expander) for each manufacturer
    for mfg in sorted(unmapped_grouped.keys()):
        with st.expander(f"📦 {mfg} Unmapped Values ({len(unmapped_grouped[mfg])})", expanded=False):
            for detail in sorted(unmapped_grouped[mfg]):
                st.write(f"- {detail}")
                
    if st.button("Acknowledge & Clear All Warnings"):
        st.session_state.unmapped_values = set()
        st.rerun()

st.divider()
st.subheader("📥 Step 1: Upload Your File to Fill")

uploaded_file = st.file_uploader("Upload your Target Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
            
        target_df.columns = target_df.columns.astype(str).str.replace('\n', ' ', regex=False).str.strip()
        
    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(f"❌ Could not find the Barcode column '{target_barcode_col}' in your file. Found columns: {list(target_df.columns)}")
        st.stop()

    st.success(f"File uploaded! Contains {len(target_df)} rows. Click below to start matching.")

    if st.button("🚀 Run Auto-Filler", type="primary"):
        with st.spinner("Matching barcodes and pouring data..."):
            
            for global_col, target_col in TARGET_MAPPING.items():
                if target_col not in target_df.columns:
                    target_df[target_col] = "" 

            match_count = 0
            
            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                
                # 🔥 ZERO-STRIPPER APPLIED HERE TOO
                clean_barcode = re.sub(r'\.0$', '', raw_barcode).lstrip('0')
                
                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]
                    
                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode": continue
                        
                        if global_col in master_db.columns:
                            val = master_row[global_col]
                            if pd.notna(val) and str(val).strip() != "":
                                target_df.at[index, target_col] = val

            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")
            
            st.write("### Preview of Filled Data:")
            st.dataframe(target_df.head(20), use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                target_df.to_excel(writer, index=False, sheet_name='Filled_Data')
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Filled Excel File",
                data=processed_data,
                file_name="Master_Filled_Glasses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )