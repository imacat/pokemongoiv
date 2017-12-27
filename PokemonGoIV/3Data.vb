' Copyright (c) 2016-2017 imacat.
' 
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
' 
'     http://www.apache.org/licenses/LICENSE-2.0
' 
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

' 3Data: The Pokémon GO data for IV calculation
'   by imacat <imacat@mail.imacat.idv.tw>, 2017-12-27
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
		Array ("Metapod", "011", 100, 45, 94, Array ("Butterfree")), _
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
		Array ("NidoranFemale", "029", 110, 86, 94, Array ("Nidorina", "Nidoqueen")), _
		Array ("Nidorina", "030", 140, 117, 126, Array ("Nidoqueen")), _
		Array ("Nidoqueen", "031", 180, 180, 174, Array ()), _
		Array ("NidoranMale", "032", 92, 105, 76, Array ("Nidorino", "Nidoking")), _
		Array ("Nidorino", "033", 122, 137, 112, Array ("Nidoking")), _
		Array ("Nidoking", "034", 162, 204, 157, Array ()), _
		Array ("Clefairy", "035", 140, 107, 116, Array ("Clefable")), _
		Array ("Clefable", "036", 190, 178, 171, Array ()), _
		Array ("Vulpix", "037", 76, 96, 122, Array ("Ninetales")), _
		Array ("Ninetales", "038", 146, 169, 204, Array ()), _
		Array ("Jigglypuff", "039", 230, 80, 44, Array ("Wigglytuff")), _
		Array ("Wigglytuff", "040", 280, 156, 93, Array ()), _
		Array ("Zubat", "041", 80, 83, 76, Array ("Golbat", "Crobat")), _
		Array ("Golbat", "042", 150, 161, 153, Array ("Crobat")), _
		Array ("Oddish", "043", 90, 131, 116, Array ("Gloom", "Vileplume", "Bellossom")), _
		Array ("Gloom", "044", 120, 153, 139, Array ("Vileplume", "Bellossom")), _
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
		Array ("Slowpoke", "079", 180, 109, 109, Array ("Slowbro", "Slowking")), _
		Array ("Slowbro", "080", 190, 177, 194, Array ()), _
		Array ("Magnemite", "081", 50, 165, 128, Array ("Magneton")), _
		Array ("Magneton", "082", 100, 223, 182, Array ()), _
		Array ("Farfetchd", "083", 104, 124, 118, Array ()), _
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
		Array ("Onix", "095", 70, 85, 288, Array ("Steelix")), _
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
		Array ("Chansey", "113", 500, 60, 176, Array ("Blissey")), _
		Array ("Tangela", "114", 130, 183, 205, Array ()), _
		Array ("Kangaskhan", "115", 210, 181, 165, Array ()), _
		Array ("Horsea", "116", 60, 129, 125, Array ("Seadra", "Kingdra")), _
		Array ("Seadra", "117", 110, 187, 182, Array ("Kingdra")), _
		Array ("Goldeen", "118", 90, 123, 115, Array ("Seaking")), _
		Array ("Seaking", "119", 160, 175, 154, Array ()), _
		Array ("Staryu", "120", 60, 137, 112, Array ("Starmie")), _
		Array ("Starmie", "121", 120, 210, 184, Array ()), _
		Array ("MrMime", "122", 80, 192, 233, Array ()), _
		Array ("Scyther", "123", 140, 218, 170, Array ("Scizor")), _
		Array ("Jynx", "124", 130, 223, 182, Array ()), _
		Array ("Electabuzz", "125", 130, 198, 173, Array ()), _
		Array ("Magmar", "126", 130, 206, 169, Array ()), _
		Array ("Pinsir", "127", 130, 238, 197, Array ()), _
		Array ("Tauros", "128", 150, 198, 197, Array ()), _
		Array ("Magikarp", "129", 40, 29, 102, Array ("Gyarados")), _
		Array ("Gyarados", "130", 190, 237, 197, Array ()), _
		Array ("Lapras", "131", 260, 165, 180, Array ()), _
		Array ("Ditto", "132", 96, 91, 91, Array ()), _
		Array ("Eevee", "133", 110, 104, 121, Array ("Vaporeon", "Jolteon", "Flareon", "Espeon", "Umbreon")), _
		Array ("Vaporeon", "134", 260, 205, 177, Array ()), _
		Array ("Jolteon", "135", 130, 232, 201, Array ()), _
		Array ("Flareon", "136", 130, 246, 204, Array ()), _
		Array ("Porygon", "137", 130, 153, 139, Array ("Porygon2")), _
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
		Array ("Mew", "151", 200, 210, 209, Array ()), _
		Array ("Chikorita", "152", 90, 92, 122, Array ("Bayleef", "Meganium")), _
		Array ("Bayleef", "153", 120, 122, 155, Array ("Meganium")), _
		Array ("Meganium", "154", 160, 168, 202, Array ()), _
		Array ("Cyndaquil", "155", 78, 116, 96, Array ("Quilava", "Typhlosion")), _
		Array ("Quilava", "156", 116, 158, 129, Array ("Typhlosion")), _
		Array ("Typhlosion", "157", 156, 223, 176, Array ()), _
		Array ("Totodile", "158", 100, 117, 116, Array ("Croconaw", "Feraligatr")), _
		Array ("Croconaw", "159", 130, 150, 151, Array ("Feraligatr")), _
		Array ("Feraligatr", "160", 170, 205, 197, Array ()), _
		Array ("Sentret", "161", 70, 79, 77, Array ("Furret")), _
		Array ("Furret", "162", 170, 148, 130, Array ()), _
		Array ("Hoothoot", "163", 120, 67, 101, Array ("Noctowl")), _
		Array ("Noctowl", "164", 200, 145, 179, Array ()), _
		Array ("Ledyba", "165", 80, 72, 142, Array ("Ledian")), _
		Array ("Ledian", "166", 110, 107, 209, Array ()), _
		Array ("Spinarak", "167", 80, 105, 73, Array ("Ariados")), _
		Array ("Ariados", "168", 140, 161, 128, Array ()), _
		Array ("Crobat", "169", 170, 194, 178, Array ()), _
		Array ("Chinchou", "170", 150, 106, 106, Array ("Lanturn")), _
		Array ("Lanturn", "171", 250, 146, 146, Array ()), _
		Array ("Pichu", "172", 40, 77, 63, Array ("Pikachu", "Raichu")), _
		Array ("Cleffa", "173", 100, 75, 91, Array ("Clefairy", "Clefable")), _
		Array ("Igglybuff", "174", 180, 69, 34, Array ("Jigglypuff", "Wigglytuff")), _
		Array ("Togepi", "175", 70, 67, 116, Array ("Togetic")), _
		Array ("Togetic", "176", 110, 139, 191, Array ()), _
		Array ("Natu", "177", 80, 134, 89, Array ("Xatu")), _
		Array ("Xatu", "178", 130, 192, 146, Array ()), _
		Array ("Mareep", "179", 110, 114, 82, Array ("Flaaffy", "Ampharos")), _
		Array ("Flaaffy", "180", 140, 145, 112, Array ("Ampharos")), _
		Array ("Ampharos", "181", 180, 211, 172, Array ()), _
		Array ("Bellossom", "182", 150, 169, 189, Array ()), _
		Array ("Marill", "183", 140, 37, 93, Array ("Azumarill")), _
		Array ("Azumarill", "184", 200, 112, 152, Array ()), _
		Array ("Sudowoodo", "185", 140, 167, 198, Array ()), _
		Array ("Politoed", "186", 180, 174, 192, Array ()), _
		Array ("Hoppip", "187", 70, 67, 101, Array ("Skiploom", "Jumpluff")), _
		Array ("Skiploom", "188", 110, 91, 127, Array ("Jumpluff")), _
		Array ("Jumpluff", "189", 150, 118, 197, Array ()), _
		Array ("Aipom", "190", 110, 136, 112, Array ()), _
		Array ("Sunkern", "191", 60, 55, 55, Array ("Sunflora")), _
		Array ("Sunflora", "192", 150, 185, 148, Array ()), _
		Array ("Yanma", "193", 130, 154, 94, Array ()), _
		Array ("Wooper", "194", 110, 75, 75, Array ("Quagsire")), _
		Array ("Quagsire", "195", 190, 152, 152, Array ()), _
		Array ("Espeon", "196", 130, 261, 194, Array ()), _
		Array ("Umbreon", "197", 190, 126, 250, Array ()), _
		Array ("Murkrow", "198", 120, 175, 87, Array ()), _
		Array ("Slowking", "199", 190, 177, 194, Array ()), _
		Array ("Misdreavus", "200", 120, 167, 167, Array ()), _
		Array ("Unown", "201", 96, 136, 91, Array ()), _
		Array ("Wobbuffet", "202", 380, 60, 106, Array ()), _
		Array ("Girafarig", "203", 140, 182, 133, Array ()), _
		Array ("Pineco", "204", 100, 108, 146, Array ("Forretress")), _
		Array ("Forretress", "205", 150, 161, 242, Array ()), _
		Array ("Dunsparce", "206", 200, 131, 131, Array ()), _
		Array ("Gligar", "207", 130, 143, 204, Array ()), _
		Array ("Steelix", "208", 150, 148, 333, Array ()), _
		Array ("Snubbull", "209", 120, 137, 89, Array ("Granbull")), _
		Array ("Granbull", "210", 180, 212, 137, Array ()), _
		Array ("Qwilfish", "211", 130, 184, 148, Array ()), _
		Array ("Scizor", "212", 140, 236, 191, Array ()), _
		Array ("Shuckle", "213", 40, 17, 396, Array ()), _
		Array ("Heracross", "214", 160, 234, 189, Array ()), _
		Array ("Sneasel", "215", 110, 189, 157, Array ()), _
		Array ("Teddiursa", "216", 120, 142, 93, Array ("Ursaring")), _
		Array ("Ursaring", "217", 180, 236, 144, Array ()), _
		Array ("Slugma", "218", 80, 118, 71, Array ("Magcargo")), _
		Array ("Magcargo", "219", 100, 139, 209, Array ()), _
		Array ("Swinub", "220", 100, 90, 74, Array ("Piloswine")), _
		Array ("Piloswine", "221", 200, 181, 147, Array ()), _
		Array ("Corsola", "222", 110, 118, 156, Array ()), _
		Array ("Remoraid", "223", 70, 127, 69, Array ("Octillery")), _
		Array ("Octillery", "224", 150, 197, 141, Array ()), _
		Array ("Delibird", "225", 90, 128, 90, Array ()), _
		Array ("Mantine", "226", 130, 148, 260, Array ()), _
		Array ("Skarmory", "227", 130, 148, 260, Array ()), _
		Array ("Houndour", "228", 90, 152, 93, Array ("Houndoom")), _
		Array ("Houndoom", "229", 150, 224, 159, Array ()), _
		Array ("Kingdra", "230", 150, 194, 194, Array ()), _
		Array ("Phanpy", "231", 180, 107, 107, Array ("Donphan")), _
		Array ("Donphan", "232", 180, 214, 214, Array ()), _
		Array ("Porygon2", "233", 170, 198, 183, Array ()), _
		Array ("Stantler", "234", 146, 192, 132, Array ()), _
		Array ("Smeargle", "235", 110, 40, 88, Array ()), _
		Array ("Tyrogue", "236", 70, 64, 64, Array ("Hitmonlee", "Hitmonchan", "Hitmontop")), _
		Array ("Hitmontop", "237", 100, 173, 214, Array ()), _
		Array ("Smoochum", "238", 90, 153, 116, Array ("Jynx")), _
		Array ("Elekid", "239", 90, 135, 110, Array ("Electabuzz")), _
		Array ("Magby", "240", 90, 151, 108, Array ("Magmar")), _
		Array ("Miltank", "241", 190, 157, 211, Array ()), _
		Array ("Blissey", "242", 510, 129, 229, Array ()), _
		Array ("Raikou", "243", 180, 241, 210, Array ()), _
		Array ("Entei", "244", 230, 235, 176, Array ()), _
		Array ("Suicune", "245", 200, 180, 235, Array ()), _
		Array ("Larvitar", "246", 100, 115, 93, Array ("Pupitar", "Tyranitar")), _
		Array ("Pupitar", "247", 140, 155, 133, Array ("Tyranitar")), _
		Array ("Tyranitar", "248", 200, 251, 212, Array ()), _
		Array ("Lugia", "249", 212, 193, 323, Array ()), _
		Array ("HoOh", "250", 193, 239, 274, Array ()), _
		Array ("Celebi", "251", 200, 210, 210, Array ()), _
		Array ("Treecko ", "252", 80, 124, 104, Array ("Grovyle ", "Sceptile ")), _
		Array ("Grovyle ", "253", 100, 172, 130, Array ("Sceptile ")), _
		Array ("Sceptile ", "254", 140, 223, 180, Array ()), _
		Array ("Torchic ", "255", 90, 130, 92, Array ("Combusken ", "Blaziken ")), _
		Array ("Combusken ", "256", 120, 163, 115, Array ("Blaziken ")), _
		Array ("Blaziken ", "257", 160, 240, 141, Array ()), _
		Array ("Mudkip ", "258", 100, 126, 93, Array ("Marshtomp ", "Swampert ")), _
		Array ("Marshtomp ", "259", 140, 156, 133, Array ("Swampert ")), _
		Array ("Swampert ", "260", 200, 208, 175, Array ()), _
		Array ("Poochyena", "261", 70, 96, 63, Array ("Mightyena")), _
		Array ("Mightyena", "262", 140, 171, 137, Array ()), _
		Array ("Zigzagoon", "263", 76, 58, 80, Array ("Linoone")), _
		Array ("Linoone", "264", 156, 142, 128, Array ()), _
		Array ("Wurmple", "265", 90, 75, 61, Array ("Silcoon", "Beautifly")), _
		Array ("Silcoon", "266", 100, 60, 91, Array ("Beautifly")), _
		Array ("Beautifly", "267", 120, 189, 98, Array ()), _
		Array ("Cascoon", "268", 100, 60, 91, Array ("Dustox")), _
		Array ("Dustox", "269", 120, 98, 172, Array ()), _
		Array ("Lotad", "270", 80, 71, 86, Array ("Lombre", "Ludicolo")), _
		Array ("Lombre", "271", 120, 112, 128, Array ("Ludicolo")), _
		Array ("Ludicolo", "272", 160, 173, 191, Array ()), _
		Array ("Seedot", "273", 80, 71, 86, Array ("Nuzleaf", "Shiftry")), _
		Array ("Nuzleaf", "274", 140, 134, 78, Array ("Shiftry")), _
		Array ("Shiftry", "275", 180, 200, 121, Array ()), _
		Array ("Taillow", "276", 80, 106, 61, Array ("Swellow")), _
		Array ("Swellow", "277", 120, 185, 130, Array ()), _
		Array ("Wingull", "278", 80, 106, 61, Array ("Pelipper")), _
		Array ("Pelipper", "279", 120, 175, 189, Array ()), _
		Array ("Ralts", "280", 56, 79, 63, Array ("Kirlia", "Gardevoir")), _
		Array ("Kirlia", "281", 76, 117, 100, Array ("Gardevoir")), _
		Array ("Gardevoir", "282", 136, 237, 220, Array ()), _
		Array ("Surskit", "283", 80, 93, 97, Array ("Masquerain")), _
		Array ("Masquerain", "284", 140, 192, 161, Array ()), _
		Array ("Shroomish", "285", 120, 74, 110, Array ("Breloom")), _
		Array ("Breloom", "286", 120, 241, 153, Array ()), _
		Array ("Slakoth", "287", 120, 104, 104, Array ("Vigoroth", "Slaking")), _
		Array ("Vigoroth", "288", 160, 159, 159, Array ("Slaking")), _
		Array ("Slaking", "289", 273, 290, 183, Array ()), _
		Array ("Nincada", "290", 62, 80, 153, Array ("Ninjask")), _
		Array ("Ninjask", "291", 122, 196, 114, Array ()), _
		Array ("Shedinja", "292", 2, 153, 80, Array ()), _
		Array ("Whismur", "293", 128, 92, 42, Array ("Loudred", "Exploud")), _
		Array ("Loudred", "294", 168, 134, 81, Array ("Exploud")), _
		Array ("Exploud", "295", 208, 179, 142, Array ()), _
		Array ("Makuhita", "296", 144, 99, 54, Array ("Hariyama")), _
		Array ("Hariyama", "297", 288, 209, 114, Array ()), _
		Array ("Azurill", "298", 100, 36, 71, Array ("Marill", "Azumarill")), _
		Array ("Nosepass", "299", 60, 82, 236, Array ()), _
		Array ("Skitty", "300", 100, 84, 84, Array ("Delcatty")), _
		Array ("Delcatty", "301", 140, 132, 132, Array ()), _
		Array ("Sableye", "302", 100, 141, 141, Array ()), _
		Array ("Mawile", "303", 100, 155, 155, Array ()), _
		Array ("Aron", "304", 100, 121, 168, Array ("Lairon", "Aggron")), _
		Array ("Lairon", "305", 120, 158, 240, Array ("Aggron")), _
		Array ("Aggron", "306", 140, 198, 314, Array ()), _
		Array ("Meditite", "307", 60, 78, 107, Array ("Medicham")), _
		Array ("Medicham", "308", 120, 121, 152, Array ()), _
		Array ("Electrike", "309", 80, 123, 78, Array ("Manectric")), _
		Array ("Manectric", "310", 140, 215, 127, Array ()), _
		Array ("Plusle", "311", 120, 167, 147, Array ()), _
		Array ("Minun", "312", 120, 147, 167, Array ()), _
		Array ("Volbeat", "313", 130, 143, 171, Array ()), _
		Array ("Illumise", "314", 130, 143, 171, Array ()), _
		Array ("Roselia", "315", 100, 186, 148, Array ()), _
		Array ("Gulpin", "316", 140, 80, 99, Array ("Swalot")), _
		Array ("Swalot", "317", 200, 140, 159, Array ()), _
		Array ("Carvanha", "318", 90, 171, 39, Array ("Sharpedo")), _
		Array ("Sharpedo", "319", 140, 243, 83, Array ()), _
		Array ("Wailmer", "320", 260, 136, 68, Array ("Wailord")), _
		Array ("Wailord", "321", 340, 175, 87, Array ()), _
		Array ("Numel", "322", 120, 119, 82, Array ("Camerupt")), _
		Array ("Camerupt", "323", 140, 194, 139, Array ()), _
		Array ("Torkoal", "324", 140, 151, 234, Array ()), _
		Array ("Spoink", "325", 120, 125, 145, Array ("Grumpig")), _
		Array ("Grumpig", "326", 160, 171, 211, Array ()), _
		Array ("Spinda", "327", 120, 116, 116, Array ()), _
		Array ("Trapinch", "328", 90, 162, 78, Array ("Vibrava", "Flygon")), _
		Array ("Vibrava", "329", 100, 134, 99, Array ("Flygon")), _
		Array ("Flygon", "330", 160, 205, 168, Array ()), _
		Array ("Cacnea", "331", 100, 156, 74, Array ("Cacturne")), _
		Array ("Cacturne", "332", 140, 221, 115, Array ()), _
		Array ("Swablu", "333", 90, 76, 139, Array ("Altaria")), _
		Array ("Altaria", "334", 150, 141, 208, Array ()), _
		Array ("Zangoose", "335", 146, 222, 124, Array ()), _
		Array ("Seviper", "336", 146, 196, 118, Array ()), _
		Array ("Lunatone", "337", 180, 178, 163, Array ()), _
		Array ("Solrock", "338", 180, 178, 163, Array ()), _
		Array ("Barboach", "339", 100, 93, 83, Array ("Whiscash")), _
		Array ("Whiscash", "340", 220, 151, 142, Array ()), _
		Array ("Corphish", "341", 86, 141, 113, Array ("Crawdaunt")), _
		Array ("Crawdaunt", "342", 126, 224, 156, Array ()), _
		Array ("Baltoy", "343", 80, 77, 131, Array ("Claydol")), _
		Array ("Claydol", "344", 120, 140, 236, Array ()), _
		Array ("Lileep", "345", 132, 105, 154, Array ("Cradily")), _
		Array ("Cradily", "346", 172, 152, 198, Array ()), _
		Array ("Anorith", "347", 90, 176, 100, Array ("Armaldo")), _
		Array ("Armaldo", "348", 150, 222, 183, Array ()), _
		Array ("Feebas", "349", 40, 29, 102, Array ("Milotic")), _
		Array ("Milotic", "350", 190, 192, 242, Array ()), _
		Array ("Castform", "351", 140, 139, 139, Array ()), _
		Array ("Kecleon", "352", 120, 161, 212, Array ()), _
		Array ("Shuppet", "353", 88, 138, 66, Array ("Banette")), _
		Array ("Banette", "354", 128, 218, 127, Array ()), _
		Array ("Duskull", "355", 40, 70, 162, Array ("Dusclops")), _
		Array ("Dusclops", "356", 80, 124, 234, Array ()), _
		Array ("Tropius", "357", 198, 136, 165, Array ()), _
		Array ("Chimecho", "358", 150, 175, 174, Array ()), _
		Array ("Absol", "359", 130, 246, 120, Array ()), _
		Array ("Wynaut", "360", 190, 41, 86, Array ("Wobbuffet")), _
		Array ("Snorunt", "361", 100, 95, 95, Array ("Glalie")), _
		Array ("Glalie", "362", 160, 162, 162, Array ()), _
		Array ("Spheal", "363", 140, 95, 90, Array ("Sealeo", "Walrein")), _
		Array ("Sealeo", "364", 180, 137, 132, Array ("Walrein")), _
		Array ("Walrein", "365", 220, 182, 176, Array ()), _
		Array ("Clamperl", "366", 70, 133, 149, Array ("Huntail")), _
		Array ("Huntail", "367", 110, 197, 194, Array ()), _
		Array ("Gorebyss", "368", 110, 211, 194, Array ()), _
		Array ("Relicanth", "369", 200, 162, 234, Array ()), _
		Array ("Luvdisc", "370", 86, 81, 134, Array ()), _
		Array ("Bagon", "371", 90, 134, 107, Array ("Shelgon", "Salamence")), _
		Array ("Shelgon", "372", 130, 172, 179, Array ("Salamence")), _
		Array ("Salamence", "373", 190, 277, 168, Array ()), _
		Array ("Beldum", "374", 80, 96, 141, Array ("Metang", "Metagross")), _
		Array ("Metang", "375", 120, 138, 185, Array ("Metagross")), _
		Array ("Metagross", "376", 160, 257, 247, Array ()), _
		Array ("Regirock", "377", 160, 179, 356, Array ()), _
		Array ("Regice", "378", 160, 179, 356, Array ()), _
		Array ("Registeel", "379", 160, 143, 285, Array ()), _
		Array ("Latias", "380", 160, 228, 268, Array ()), _
		Array ("Latios", "381", 160, 268, 228, Array ()), _
		Array ("Kyogre", "382", 182, 270, 251, Array ()), _
		Array ("Groudon", "383", 182, 270, 251, Array ()), _
		Array ("Rayquaza", "384", 191, 284, 170, Array ()), _
		Array ("Jirachi", "385", 200, 210, 210, Array ()), _
		Array ("Deoxys", "386", 1, 1, 1, Array ()))
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

' fnGetStardustData: Returns the stardust data.
Function fnGetStardustData As Variant
	fnGetStardustData = Array( _
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
