' 1Data: The Pokémon GO data for IV calculation
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-12-06
'   Generated with 9Load.subReadDataSheets ()

Option Explicit

' fnGetBaseStatsData: Returns the base stats data.
Function fnGetBaseStatsData As Variant
	fnGetBaseStatsData = Array( _
		Array ("Bulbasaur", "001", 90, 118, 118, Array ("Ivysaur", "Venusaur")), _
		Array ("Ivysaur", "002", 120, 151, 151, Array ("Venusaur")), _
		Array ("Venusaur", "003", 160, 198, 198, Array ()), _
		Array ("Charmander", "004", 78, 116, 96, Array ("Charmeleon", "Charizard")), _
		Array ("Charmeleon", "005", 116, 158, 129, Array ("Charizard")), _
		Array ("Charizard", "006", 156, 223, 176, Array ()), _
		Array ("Squirtle", "007", 88, 94, 122, Array ("Wartortle", "Blastoise")), _
		Array ("Wartortle", "008", 118, 126, 155, Array ("Blastoise")), _
		Array ("Blastoise", "009", 158, 171, 210, Array ()), _
		Array ("Caterpie", "010", 90, 55, 62, Array ("Metapod", "Butterfree")), _
		Array ("Metapod", "011", 100, 45, 64, Array ("Butterfree")), _
		Array ("Butterfree", "012", 120, 167, 151, Array ()), _
		Array ("Weedle", "013", 80, 63, 55, Array ("Kakuna", "Beedrill")), _
		Array ("Kakuna", "014", 90, 46, 86, Array ("Beedrill")), _
		Array ("Beedrill", "015", 130, 169, 150, Array ()), _
		Array ("Pidgey", "016", 80, 85, 76, Array ("Pidgeotto", "Pidgeot")), _
		Array ("Pidgeotto", "017", 126, 117, 108, Array ("Pidgeot")), _
		Array ("Pidgeot", "018", 166, 166, 157, Array ()), _
		Array ("Rattata", "019", 60, 103, 70, Array ("Raticate")), _
		Array ("Raticate", "020", 110, 161, 144, Array ()), _
		Array ("Spearow", "021", 80, 112, 61, Array ("Fearow")), _
		Array ("Fearow", "022", 130, 182, 135, Array ()), _
		Array ("Ekans", "023", 70, 110, 102, Array ("Arbok")), _
		Array ("Arbok", "024", 120, 167, 158, Array ()), _
		Array ("Pikachu", "025", 70, 112, 101, Array ("Raichu")), _
		Array ("Raichu", "026", 120, 193, 165, Array ()), _
		Array ("Sandshrew", "027", 100, 126, 145, Array ("Sandslash")), _
		Array ("Sandslash", "028", 150, 182, 202, Array ()), _
		Array ("Nidoran♀", "029", 110, 86, 94, Array ("Nidorina", "Nidoqueen")), _
		Array ("Nidorina", "030", 140, 117, 126, Array ("Nidoqueen")), _
		Array ("Nidoqueen", "031", 180, 180, 174, Array ()), _
		Array ("Nidoran♂", "032", 92, 105, 76, Array ("Nidorino", "Nidoking")), _
		Array ("Nidorino", "033", 122, 137, 112, Array ("Nidoking")), _
		Array ("Nidoking", "034", 162, 204, 157, Array ()), _
		Array ("Clefairy", "035", 140, 107, 116, Array ("Clefable")), _
		Array ("Clefable", "036", 190, 178, 171, Array ()), _
		Array ("Vulpix", "037", 76, 96, 122, Array ("Ninetales")), _
		Array ("Ninetales", "038", 146, 169, 204, Array ()), _
		Array ("Jigglypuff", "039", 230, 80, 44, Array ("Wigglytuff")), _
		Array ("Wigglytuff", "040", 280, 156, 93, Array ()), _
		Array ("Zubat", "041", 80, 83, 76, Array ("Golbat")), _
		Array ("Golbat", "042", 150, 161, 153, Array ()), _
		Array ("Oddish", "043", 90, 131, 116, Array ("Gloom", "Vileplume")), _
		Array ("Gloom", "044", 120, 153, 139, Array ("Vileplume")), _
		Array ("Vileplume", "045", 150, 202, 170, Array ()), _
		Array ("Paras", "046", 70, 121, 99, Array ("Parasect")), _
		Array ("Parasect", "047", 120, 165, 146, Array ()), _
		Array ("Venonat", "048", 120, 100, 102, Array ("Venomoth")), _
		Array ("Venomoth", "049", 140, 179, 150, Array ()), _
		Array ("Diglett", "050", 20, 109, 88, Array ("Dugtrio")), _
		Array ("Dugtrio", "051", 70, 167, 147, Array ()), _
		Array ("Meowth", "052", 80, 92, 81, Array ("Persian")), _
		Array ("Persian", "053", 130, 150, 139, Array ()), _
		Array ("Psyduck", "054", 100, 122, 96, Array ("Golduck")), _
		Array ("Golduck", "055", 160, 191, 163, Array ()), _
		Array ("Mankey", "056", 80, 148, 87, Array ("Primeape")), _
		Array ("Primeape", "057", 130, 207, 144, Array ()), _
		Array ("Growlithe", "058", 110, 136, 96, Array ("Arcanine")), _
		Array ("Arcanine", "059", 180, 227, 166, Array ()), _
		Array ("Poliwag", "060", 80, 101, 82, Array ("Poliwhirl", "Poliwrath")), _
		Array ("Poliwhirl", "061", 130, 130, 130, Array ("Poliwrath")), _
		Array ("Poliwrath", "062", 180, 182, 187, Array ()), _
		Array ("Abra", "063", 50, 195, 103, Array ("Kadabra", "Alakazam")), _
		Array ("Kadabra", "064", 80, 232, 138, Array ("Alakazam")), _
		Array ("Alakazam", "065", 110, 271, 194, Array ()), _
		Array ("Machop", "066", 140, 137, 88, Array ("Machoke", "Machamp")), _
		Array ("Machoke", "067", 160, 177, 130, Array ("Machamp")), _
		Array ("Machamp", "068", 180, 234, 162, Array ()), _
		Array ("Bellsprout", "069", 100, 139, 64, Array ("Weepinbell", "Victreebel")), _
		Array ("Weepinbell", "070", 130, 172, 95, Array ("Victreebel")), _
		Array ("Victreebel", "071", 160, 207, 138, Array ()), _
		Array ("Tentacool", "072", 80, 97, 182, Array ("Tentacruel")), _
		Array ("Tentacruel", "073", 160, 166, 237, Array ()), _
		Array ("Geodude", "074", 80, 132, 163, Array ("Graveler", "Golem")), _
		Array ("Graveler", "075", 110, 164, 196, Array ("Golem")), _
		Array ("Golem", "076", 160, 211, 229, Array ()), _
		Array ("Ponyta", "077", 100, 170, 132, Array ("Rapidash")), _
		Array ("Rapidash", "078", 130, 207, 167, Array ()), _
		Array ("Slowpoke", "079", 180, 109, 109, Array ("Slowbro")), _
		Array ("Slowbro", "080", 190, 177, 194, Array ()), _
		Array ("Magnemite", "081", 50, 165, 128, Array ("Magneton")), _
		Array ("Magneton", "082", 100, 223, 182, Array ()), _
		Array ("Farfetch'd", "083", 104, 124, 118, Array ()), _
		Array ("Doduo", "084", 70, 158, 88, Array ("Dodrio")), _
		Array ("Dodrio", "085", 120, 218, 145, Array ()), _
		Array ("Seel", "086", 130, 85, 128, Array ("Dewgong")), _
		Array ("Dewgong", "087", 180, 139, 184, Array ()), _
		Array ("Grimer", "088", 160, 135, 90, Array ("Muk")), _
		Array ("Muk", "089", 210, 190, 184, Array ()), _
		Array ("Shellder", "090", 60, 116, 168, Array ("Cloyster")), _
		Array ("Cloyster", "091", 100, 186, 323, Array ()), _
		Array ("Gastly", "092", 60, 186, 70, Array ("Haunter", "Gengar")), _
		Array ("Haunter", "093", 90, 223, 112, Array ("Gengar")), _
		Array ("Gengar", "094", 120, 261, 156, Array ()), _
		Array ("Onix", "095", 70, 85, 288, Array ()), _
		Array ("Drowzee", "096", 120, 89, 158, Array ("Hypno")), _
		Array ("Hypno", "097", 170, 144, 215, Array ()), _
		Array ("Krabby", "098", 60, 181, 156, Array ("Kingler")), _
		Array ("Kingler", "099", 110, 240, 214, Array ()), _
		Array ("Voltorb", "100", 80, 109, 114, Array ("Electrode")), _
		Array ("Electrode", "101", 120, 173, 179, Array ()), _
		Array ("Exeggcute", "102", 120, 107, 140, Array ("Exeggutor")), _
		Array ("Exeggutor", "103", 190, 233, 158, Array ()), _
		Array ("Cubone", "104", 100, 90, 165, Array ("Marowak")), _
		Array ("Marowak", "105", 120, 144, 200, Array ()), _
		Array ("Hitmonlee", "106", 100, 224, 211, Array ()), _
		Array ("Hitmonchan", "107", 100, 193, 212, Array ()), _
		Array ("Lickitung", "108", 180, 108, 137, Array ()), _
		Array ("Koffing", "109", 80, 119, 164, Array ("Weezing")), _
		Array ("Weezing", "110", 130, 174, 221, Array ()), _
		Array ("Rhyhorn", "111", 160, 140, 157, Array ("Rhydon")), _
		Array ("Rhydon", "112", 210, 222, 206, Array ()), _
		Array ("Chansey", "113", 500, 60, 176, Array ()), _
		Array ("Tangela", "114", 130, 183, 205, Array ()), _
		Array ("Kangaskhan", "115", 210, 181, 165, Array ()), _
		Array ("Horsea", "116", 60, 129, 125, Array ("Seadra")), _
		Array ("Seadra", "117", 110, 187, 182, Array ()), _
		Array ("Goldeen", "118", 90, 123, 115, Array ("Seaking")), _
		Array ("Seaking", "119", 160, 175, 154, Array ()), _
		Array ("Staryu", "120", 60, 137, 112, Array ("Starmie")), _
		Array ("Starmie", "121", 120, 210, 184, Array ()), _
		Array ("Mr. Mime", "122", 80, 192, 233, Array ()), _
		Array ("Scyther", "123", 140, 218, 170, Array ()), _
		Array ("Jynx", "124", 130, 223, 182, Array ()), _
		Array ("Electabuzz", "125", 130, 198, 173, Array ()), _
		Array ("Magmar", "126", 130, 206, 169, Array ()), _
		Array ("Pinsir", "127", 130, 238, 197, Array ()), _
		Array ("Tauros", "128", 150, 198, 197, Array ()), _
		Array ("Magikarp", "129", 40, 29, 102, Array ("Gyarados")), _
		Array ("Gyarados", "130", 190, 237, 197, Array ()), _
		Array ("Lapras", "131", 260, 186, 190, Array ()), _
		Array ("Ditto", "132", 96, 91, 91, Array ()), _
		Array ("Eevee", "133", 110, 104, 121, Array ("Vaporeon", "Jolteon", "Flareon")), _
		Array ("Vaporeon", "134", 260, 205, 177, Array ()), _
		Array ("Jolteon", "135", 130, 232, 201, Array ()), _
		Array ("Flareon", "136", 130, 246, 204, Array ()), _
		Array ("Porygon", "137", 130, 153, 139, Array ()), _
		Array ("Omanyte", "138", 70, 155, 174, Array ("Omastar")), _
		Array ("Omastar", "139", 140, 207, 227, Array ()), _
		Array ("Kabuto", "140", 60, 148, 162, Array ("Kabutops")), _
		Array ("Kabutops", "141", 120, 220, 203, Array ()), _
		Array ("Aerodactyl", "142", 160, 221, 164, Array ()), _
		Array ("Snorlax", "143", 320, 190, 190, Array ()), _
		Array ("Articuno", "144", 180, 192, 249, Array ()), _
		Array ("Zapdos", "145", 180, 253, 188, Array ()), _
		Array ("Moltres", "146", 180, 251, 184, Array ()), _
		Array ("Dratini", "147", 82, 119, 94, Array ("Dragonair", "Dragonite")), _
		Array ("Dragonair", "148", 122, 163, 138, Array ("Dragonite")), _
		Array ("Dragonite", "149", 182, 263, 201, Array ()), _
		Array ("Mewtwo", "150", 212, 330, 200, Array ()), _
		Array ("Mew", "151", 200, 210, 209, Array ()))
End Function

' fnGetCPMData: Returns the combat power multiplier data.
Function fnGetCPMData As Variant
	fnGetCPMData = Array( _
		-1, _
		9.4E-02, _
		0.16639787, _
		0.21573247, _
		0.25572005, _
		0.29024988, _
		0.3210876, _
		0.34921268, _
		0.37523559, _
		0.39956728, _
		0.42250001, _
		0.44310755, _
		0.46279839, _
		0.48168495, _
		0.49985844, _
		0.51739395, _
		0.53435433, _
		0.55079269, _
		0.56675452, _
		0.58227891, _
		0.59740001, _
		0.61215729, _
		0.62656713, _
		0.64065295, _
		0.65443563, _
		0.667934, _
		0.68116492, _
		0.69414365, _
		0.70688421, _
		0.71939909, _
		0.7317, _
		0.73776948, _
		0.74378943, _
		0.74976104, _
		0.75568551, _
		0.76156384, _
		0.76739717, _
		0.7731865, _
		0.77893275, _
		0.78463697, _
		0.78463697)
End Function

' fnGetStarDustData: Returns the star dust data.
Function fnGetStarDustData As Variant
	fnGetStarDustData = Array( _
		-1, _
		200, _
		200, _
		400, _
		400, _
		600, _
		600, _
		800, _
		800, _
		1000, _
		1000, _
		1300, _
		1300, _
		1600, _
		1600, _
		1900, _
		1900, _
		2200, _
		2200, _
		2500, _
		2500, _
		3000, _
		3000, _
		3500, _
		3500, _
		4000, _
		4000, _
		4500, _
		4500, _
		5000, _
		5000, _
		6000, _
		6000, _
		7000, _
		7000, _
		8000, _
		8000, _
		9000, _
		9000, _
		10000, _
		10000)
End Function
