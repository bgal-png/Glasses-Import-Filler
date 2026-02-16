# mappings.py
# This file contains the data lists for the Excel Auto-Filler.

# --- RULE 3: PRODUCING COMPANY ---
BRAND_TO_COMPANY_MAP = {
    "Kering": ["Alexander McQueen", "Balenciaga", "Chloe", "Gucci", "Maui Jim", "Montblanc", "Puma", "Saint Laurent"],
    "Marcolin": ["Adidas", "Guess", "Max Mara", "MAX&Co.", "Tom Ford"],
    "Ostalo": ["Arena", "Cebe", "Hawkers", "HEAD", "Lavida", "POC", "Oxydo", "Alpina'", "Alpina"],
    "Inspecs": ["Caterpillar", "O'Neill", "Radley", "Superdry"],
    "Marchon": ["Calvin Klein", "Lacoste", "LIU JO", "Nike"],
    "Alensa": ["Alensa"],
    "Adrial": ["Crullé", "Kimikado", "Marisio", "Válle", "LeWish", "Beron"],
    "Luxottica": ["Arnette", "Burberry", "Dolce & Gabbana", "Emporio Armani", "Giorgio Armani", "Armani Exchange", "Michael Kors", "Oakley", "Persol", "Polo Ralph Lauren", "Prada", "Ralph by Ralph Lauren", "Ray-Ban", "Swarovski", "Versace", "Vogue", "Jimmy Choo", "Miu Miu", "Tiffany", "Ralph Lauren"],
    "Safilo": ["Boss by Hugo Boss", "Carrera", "David Beckham", "Love Moschino", "Chiara Ferragni", "Dsquared2", "Fossil", "Havaianas", "Hugo by Hugo Boss", "Kate Spade", "Levi's", "Marc Jacobs", "Missoni", "Moschino", "Pierre Cardin", "Polaroid", "Tommy Hilfiger", "Under Armour", "Seventh Street", "Carolina Herrera"],
    "GO Eyewear": ["Ana Hickmann"],
    "Strabilia": ["Silhouette"],
    "MCM OPTIK SRL": ["Morel"],
    "Bollé Brands": ["Bollé", "SPY+", "Serengeti"],
    "": [] 
}

# --- RULE 4: GLASSES USABLE ---
BRAND_TO_USABLE_MAP = {
    "Fashion glasses": ["Botaniq", "Brioni", "Calvin Klein", "Carrera", "Coco song", "Crullé", "Dsquared2", "Fossil", "Guess", "Hawkers", "Hugo Boss", "BOSS", "Boss by Hugo Boss", "Julbo", "Kate Spade", "Kimikado", "Lacoste", "Levis", "Levi's", "Marc Jacobs", "Marisio", "Max Mara", "Max&Co.", "Michael Kors", "Persol", "Polaroid", "Police", "Puma", "Radley", "Ray-Ban", "Seventh Street", "Superdry", "Swarovski", "Swidoo", "Tommy Hilfiger", "Vogue", "Under Armour", "Armani Exchange", "Nike"],
    "Luxury glasses": ["Alexander McQueen", "Balenciaga", "Bottega Venetta", "Burberry", "Celine", "Chiara Ferreagni", "Chloe", "Christian Dior", "Dolce & Gabbana", "Emporio Armani", "Fendi", "Givenchy", "Gucci", "Impressio", "Jimmy Choo", "Liu Jo", "Michael Kors", "Missoni", "Moschino", "Love Moschino", "Myth", "Pierre Cardin", "Polo Ralph Lauren", "Ralph by Ralph Lauren", "Prada", "Saint Laurent", "Stella McCarteny", "Tiffany", "Tom Ford", "Versace", "Miu Miu", "Beron", "LeWish", "Giorgio Armani", "Carolina Herrera", "David Beckham", "Ralph Lauren", "Victoria Beckham"]
}

# --- RULE 5: GLASSES COLLECTION ---
BRAND_TO_COLLECTION_MAP = {
    "Prémiové brýle - Kering": ["Alexander McQueen", "Balenciaga", "Gucci", "Saint Laurent", "Chloe", "Dior", "Christian Dior", "Fendi", "Dolce & Gabbana", "Celine", "Miu Miu", "Tom Ford", "Prada", "Giorgio Armani", "Beron", "LeWish"]
}

# --- RULE 8: FACE SHAPE RECOMMENDATION ---
FACE_SHAPE_MAP = {
    "Round face|Oval face|Heart-shaped face|Square face": ["Extravagant", "Pilot", "Oversize", "Panthos / Tea cup", "Single lens", "Browline"],
    "Oval face|Square face": ["Round", "Butterfly"],
    "Oval face|Heart-shaped face": ["Cat Eye"],
    "Round face|Oval face|Heart-shaped face": ["Square", "Rectangular"],
    "Oval face|Heart-shaped face|Square face": ["Oval / Elipse"],
    "Round face|Oval face": ["Hexagonal"]
}