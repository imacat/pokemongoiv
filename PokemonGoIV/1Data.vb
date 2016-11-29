' 1Data: The Pokémon Go data for IV calculation
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-11-28
'   Generated with 9Load.subReadDataSheets ()

Option Explicit

' fnGetBaseStatsData: Returns the base stats data.
Function fnGetBaseStatsData As Variant
	fnGetBaseStatsData = Array( _
		Array ("Bulbasaur", "001", 90, 118, 118, "Venusaur"), _
		Array ("Ivysaur", "002", 120, 151, 151, "Venusaur"), _
		Array ("Venusaur", "003", 160, 198, 198, "Venusaur"), _
		Array ("Charmander", "004", 78, 116, 96, "Charizard"), _
		Array ("Charmeleon", "005", 116, 158, 129, "Charizard"), _
		Array ("Charizard", "006", 156, 223, 176, "Charizard"), _
		Array ("Squirtle", "007", 88, 94, 122, "Blastoise"), _
		Array ("Wartortle", "008", 118, 126, 155, "Blastoise"), _
		Array ("Blastoise", "009", 158, 171, 210, "Blastoise"), _
		Array ("Caterpie", "010", 90, 55, 62, "Butterfree"), _
		Array ("Metapod", "011", 100, 45, 64, "Butterfree"), _
		Array ("Butterfree", "012", 120, 167, 151, "Butterfree"), _
		Array ("Weedle", "013", 80, 63, 55, "Beedrill"), _
		Array ("Kakuna", "014", 90, 46, 86, "Beedrill"), _
		Array ("Beedrill", "015", 130, 169, 150, "Beedrill"), _
		Array ("Pidgey", "016", 80, 85, 76, "Pidgeot"), _
		Array ("Pidgeotto", "017", 126, 117, 108, "Pidgeot"), _
		Array ("Pidgeot", "018", 166, 166, 157, "Pidgeot"), _
		Array ("Rattata", "019", 60, 103, 70, "Raticate"), _
		Array ("Raticate", "020", 110, 161, 144, "Raticate"), _
		Array ("Spearow", "021", 80, 112, 61, "Fearow"), _
		Array ("Fearow", "022", 130, 182, 135, "Fearow"), _
		Array ("Ekans", "023", 70, 110, 102, "Arbok"), _
		Array ("Arbok", "024", 120, 167, 158, "Arbok"), _
		Array ("Pikachu", "025", 70, 112, 101, "Raichu"), _
		Array ("Raichu", "026", 120, 193, 165, "Raichu"), _
		Array ("Sandshrew", "027", 100, 126, 145, "Sandslash"), _
		Array ("Sandslash", "028", 150, 182, 202, "Sandslash"), _
		Array ("Nidoran♀", "029", 110, 86, 94, "Nidoqueen"), _
		Array ("Nidorina", "030", 140, 117, 126, "Nidoqueen"), _
		Array ("Nidoqueen", "031", 180, 180, 174, "Nidoqueen"), _
		Array ("Nidoran♂", "032", 92, 105, 76, "Nidoking"), _
		Array ("Nidorino", "033", 122, 137, 112, "Nidoking"), _
		Array ("Nidoking", "034", 162, 204, 157, "Nidoking"), _
		Array ("Clefairy", "035", 140, 107, 116, "Clefable"), _
		Array ("Clefable", "036", 190, 178, 171, "Clefable"), _
		Array ("Vulpix", "037", 76, 96, 122, "Ninetales"), _
		Array ("Ninetales", "038", 146, 169, 204, "Ninetales"), _
		Array ("Jigglypuff", "039", 230, 80, 44, "Wigglytuff"), _
		Array ("Wigglytuff", "040", 280, 156, 93, "Wigglytuff"), _
		Array ("Zubat", "041", 80, 83, 76, "Golbat"), _
		Array ("Golbat", "042", 150, 161, 153, "Golbat"), _
		Array ("Oddish", "043", 90, 131, 116, "Vileplume"), _
		Array ("Gloom", "044", 120, 153, 139, "Vileplume"), _
		Array ("Vileplume", "045", 150, 202, 170, "Vileplume"), _
		Array ("Paras", "046", 70, 121, 99, "Parasect"), _
		Array ("Parasect", "047", 120, 165, 146, "Parasect"), _
		Array ("Venonat", "048", 120, 100, 102, "Venomoth"), _
		Array ("Venomoth", "049", 140, 179, 150, "Venomoth"), _
		Array ("Diglett", "050", 20, 109, 88, "Dugtrio"), _
		Array ("Dugtrio", "051", 70, 167, 147, "Dugtrio"), _
		Array ("Meowth", "052", 80, 92, 81, "Persian"), _
		Array ("Persian", "053", 130, 150, 139, "Persian"), _
		Array ("Psyduck", "054", 100, 122, 96, "Golduck"), _
		Array ("Golduck", "055", 160, 191, 163, "Golduck"), _
		Array ("Mankey", "056", 80, 148, 87, "Primeape"), _
		Array ("Primeape", "057", 130, 207, 144, "Primeape"), _
		Array ("Growlithe", "058", 110, 136, 96, "Arcanine"), _
		Array ("Arcanine", "059", 180, 227, 166, "Arcanine"), _
		Array ("Poliwag", "060", 80, 101, 82, "Poliwrath"), _
		Array ("Poliwhirl", "061", 130, 130, 130, "Poliwrath"), _
		Array ("Poliwrath", "062", 180, 182, 187, "Poliwrath"), _
		Array ("Abra", "063", 50, 195, 103, "Alakazam"), _
		Array ("Kadabra", "064", 80, 232, 138, "Alakazam"), _
		Array ("Alakazam", "065", 110, 271, 194, "Alakazam"), _
		Array ("Machop", "066", 140, 137, 88, "Machamp"), _
		Array ("Machoke", "067", 160, 177, 130, "Machamp"), _
		Array ("Machamp", "068", 180, 234, 162, "Machamp"), _
		Array ("Bellsprout", "069", 100, 139, 64, "Victreebel"), _
		Array ("Weepinbell", "070", 130, 172, 95, "Victreebel"), _
		Array ("Victreebel", "071", 160, 207, 138, "Victreebel"), _
		Array ("Tentacool", "072", 80, 97, 182, "Tentacruel"), _
		Array ("Tentacruel", "073", 160, 166, 237, "Tentacruel"), _
		Array ("Geodude", "074", 80, 132, 163, "Golem"), _
		Array ("Graveler", "075", 110, 164, 196, "Golem"), _
		Array ("Golem", "076", 160, 211, 229, "Golem"), _
		Array ("Ponyta", "077", 100, 170, 132, "Rapidash"), _
		Array ("Rapidash", "078", 130, 207, 167, "Rapidash"), _
		Array ("Slowpoke", "079", 180, 109, 109, "Slowbro"), _
		Array ("Slowbro", "080", 190, 177, 194, "Slowbro"), _
		Array ("Magnemite", "081", 50, 165, 128, "Magneton"), _
		Array ("Magneton", "082", 100, 223, 182, "Magneton"), _
		Array ("Farfetch'd", "083", 104, 124, 118, "Farfetch'd"), _
		Array ("Doduo", "084", 70, 158, 88, "Dodrio"), _
		Array ("Dodrio", "085", 120, 218, 145, "Dodrio"), _
		Array ("Seel", "086", 130, 85, 128, "Dewgong"), _
		Array ("Dewgong", "087", 180, 139, 184, "Dewgong"), _
		Array ("Grimer", "088", 160, 135, 90, "Muk"), _
		Array ("Muk", "089", 210, 190, 184, "Muk"), _
		Array ("Shellder", "090", 60, 116, 168, "Cloyster"), _
		Array ("Cloyster", "091", 100, 186, 323, "Cloyster"), _
		Array ("Gastly", "092", 60, 186, 70, "Gengar"), _
		Array ("Haunter", "093", 90, 223, 112, "Gengar"), _
		Array ("Gengar", "094", 120, 261, 156, "Gengar"), _
		Array ("Onix", "095", 70, 85, 288, "Onix"), _
		Array ("Drowzee", "096", 120, 89, 158, "Hypno"), _
		Array ("Hypno", "097", 170, 144, 215, "Hypno"), _
		Array ("Krabby", "098", 60, 181, 156, "Kingler"), _
		Array ("Kingler", "099", 110, 240, 214, "Kingler"), _
		Array ("Voltorb", "100", 80, 109, 114, "Electrode"), _
		Array ("Electrode", "101", 120, 173, 179, "Electrode"), _
		Array ("Exeggcute", "102", 120, 107, 140, "Exeggutor"), _
		Array ("Exeggutor", "103", 190, 233, 158, "Exeggutor"), _
		Array ("Cubone", "104", 100, 90, 165, "Marowak"), _
		Array ("Marowak", "105", 120, 144, 200, "Marowak"), _
		Array ("Hitmonlee", "106", 100, 224, 211, "Hitmonlee"), _
		Array ("Hitmonchan", "107", 100, 193, 212, "Hitmonchan"), _
		Array ("Lickitung", "108", 180, 108, 137, "Lickitung"), _
		Array ("Koffing", "109", 80, 119, 164, "Weezing"), _
		Array ("Weezing", "110", 130, 174, 221, "Weezing"), _
		Array ("Rhyhorn", "111", 160, 140, 157, "Rhydon"), _
		Array ("Rhydon", "112", 210, 222, 206, "Rhydon"), _
		Array ("Chansey", "113", 500, 60, 176, "Chansey"), _
		Array ("Tangela", "114", 130, 183, 205, "Tangela"), _
		Array ("Kangaskhan", "115", 210, 181, 165, "Kangaskhan"), _
		Array ("Horsea", "116", 60, 129, 125, "Seadra"), _
		Array ("Seadra", "117", 110, 187, 182, "Seadra"), _
		Array ("Goldeen", "118", 90, 123, 115, "Seaking"), _
		Array ("Seaking", "119", 160, 175, 154, "Seaking"), _
		Array ("Staryu", "120", 60, 137, 112, "Starmie"), _
		Array ("Starmie", "121", 120, 210, 184, "Starmie"), _
		Array ("Mr. Mime", "122", 80, 192, 233, "Mr. Mime"), _
		Array ("Scyther", "123", 140, 218, 170, "Scyther"), _
		Array ("Jynx", "124", 130, 223, 182, "Jynx"), _
		Array ("Electabuzz", "125", 130, 198, 173, "Electabuzz"), _
		Array ("Magmar", "126", 130, 206, 169, "Magmar"), _
		Array ("Pinsir", "127", 130, 238, 197, "Pinsir"), _
		Array ("Tauros", "128", 150, 198, 197, "Tauros"), _
		Array ("Magikarp", "129", 40, 29, 102, "Gyarados"), _
		Array ("Gyarados", "130", 190, 237, 197, "Gyarados"), _
		Array ("Lapras", "131", 260, 186, 190, "Lapras"), _
		Array ("Ditto", "132", 96, 91, 91, "Ditto"), _
		Array ("Eevee", "133", 110, 104, 121, "Vaporeon"), _
		Array ("Vaporeon", "134", 260, 205, 177, "Vaporeon"), _
		Array ("Jolteon", "135", 130, 232, 201, "Jolteon"), _
		Array ("Flareon", "136", 130, 246, 204, "Flareon"), _
		Array ("Porygon", "137", 130, 153, 139, "Porygon"), _
		Array ("Omanyte", "138", 70, 155, 174, "Omastar"), _
		Array ("Omastar", "139", 140, 207, 227, "Omastar"), _
		Array ("Kabuto", "140", 60, 148, 162, "Kabutops"), _
		Array ("Kabutops", "141", 120, 220, 203, "Kabutops"), _
		Array ("Aerodactyl", "142", 160, 221, 164, "Aerodactyl"), _
		Array ("Snorlax", "143", 320, 190, 190, "Snorlax"), _
		Array ("Articuno", "144", 180, 192, 249, "Articuno"), _
		Array ("Zapdos", "145", 180, 253, 188, "Zapdos"), _
		Array ("Moltres", "146", 180, 251, 184, "Moltres"), _
		Array ("Dratini", "147", 82, 119, 94, "Dragonite"), _
		Array ("Dragonair", "148", 122, 163, 138, "Dragonite"), _
		Array ("Dragonite", "149", 182, 263, 201, "Dragonite"), _
		Array ("Mewtwo", "150", 212, 330, 200, "Mewtwo"), _
		Array ("Mew", "151", 200, 210, 209, "Mew"))
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
