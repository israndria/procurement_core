# ======================================================
# MAGIC CHESS: GO GO - MASTER DATABASE (USER DEFINED)
# Updated: 5 Januari 2026 (Revision: Tank -> Defender)
# ======================================================

# ------------------------------------------------------
# BAGIAN 1: DATABASE HERO (Role, Faksi, & Harga)
# ------------------------------------------------------
hero_db = {
    # --- SOUL VESSELS ---
    "Cecilion":  {"syn": ["Soul Vessels", "Mage"],        "cost": 1},
    "Clint":     {"syn": ["Soul Vessels", "Phasewarper"], "cost": 2},
    "Aamon":     {"syn": ["Soul Vessels", "Assassin"],    "cost": 2},
    "Gloo":      {"syn": ["Soul Vessels", "Dauntless"],   "cost": 3},
    "Hanabi":    {"syn": ["Soul Vessels", "Marksman"],    "cost": 4},
    "Benedetta": {"syn": ["Soul Vessels", "Fighter"],     "cost": 5},

    # --- ASPIRANTS ---
    "Layla":     {"syn": ["Aspirants", "Marksman"],          "cost": 1},
    "Ruby":      {"syn": ["Aspirants", "Dauntless", "Defender"], "cost": 2}, # Unik: 3 Sinergi (Tank -> Defender)
    "Fanny":     {"syn": ["Aspirants", "Assassin"],          "cost": 3},
    "Lesley":    {"syn": ["Aspirants", "Phasewarper"],       "cost": 3},
    "Guinevere": {"syn": ["Aspirants", "Bruiser"],           "cost": 4},
    "Vexana":    {"syn": ["Aspirants", "Stargazer"],         "cost": 5},

    # --- SHADOWCELL ---
    "Alucard":   {"syn": ["Shadowcell", "Fighter"],     "cost": 1},
    "Helcurt":   {"syn": ["Shadowcell", "Assassin"],    "cost": 1},
    "Obsidia":   {"syn": ["Shadowcell", "Marksman"],    "cost": 2},
    "Masha":     {"syn": ["Shadowcell", "Bruiser"],     "cost": 3},
    "Valentina": {"syn": ["Shadowcell", "Mage"],        "cost": 4},
    "Arlott":    {"syn": ["Shadowcell", "Phasewarper"], "cost": 4},

    # --- STARWING ---
    "Atlas":    {"syn": ["Starwing", "Defender"],    "cost": 1}, # Tank -> Defender
    "Mathilda": {"syn": ["Starwing", "Stargazer"],   "cost": 2},
    "Belerick": {"syn": ["Starwing", "Bruiser"],     "cost": 3},
    "Saber":    {"syn": ["Starwing", "Assassin"],    "cost": 4},
    "Freya":    {"syn": ["Starwing", "Fighter"],     "cost": 4},
    "Harley":   {"syn": ["Starwing", "Phasewarper"], "cost": 5},

    # --- LUMINEXUS ---
    "Lolita":   {"syn": ["Luminexus", "Dauntless"], "cost": 1},
    "Carmilla": {"syn": ["Luminexus", "Scavenger"], "cost": 2},
    "Cici":     {"syn": ["Luminexus", "Fighter"],   "cost": 2},
    "Zhuxin":   {"syn": ["Luminexus", "Mage"],      "cost": 3},
    "Kadita":   {"syn": ["Luminexus", "Stargazer"], "cost": 4},
    "Irithel":  {"syn": ["Luminexus", "Marksman"],  "cost": 5},

    # --- TOY MISCHIEF ---
    "Jawhead": {"syn": ["Toy Mischief", "Bruiser"],   "cost": 1},
    "Uranus":  {"syn": ["Toy Mischief", "Defender"],  "cost": 2}, # Tank -> Defender
    "Harith":  {"syn": ["Toy Mischief", "Mage"],      "cost": 3},
    "Aulus":   {"syn": ["Toy Mischief", "Fighter"],   "cost": 3},
    "Barats":  {"syn": ["Toy Mischief", "Dauntless"], "cost": 4},
    "Akai":    {"syn": ["Toy Mischief", "Scavenger"], "cost": 5},

    # --- BEYOND THE CLOUDS ---
    "Kagura": {"syn": ["Beyond the Clouds", "Mage"],      "cost": 2},
    "Xavier": {"syn": ["Beyond the Clouds", "Stargazer"], "cost": 3},
    "Edith":  {"syn": ["Beyond the Clouds", "Defender"],  "cost": 5}, # Tank -> Defender

    # --- K.O.F & MORTAL RIVAL ---
    "Kula":         {"syn": ["KOF", "Stargazer"], "cost": 1},
    "Chris":        {"syn": ["KOF", "Dauntless"], "cost": 2},
    "Leona":        {"syn": ["KOF", "Scavenger"], "cost": 3},
    "K'":           {"syn": ["KOF", "Assassin"],  "cost": 4},
    "Terry Bogard": {"syn": ["KOF", "Defender"],  "cost": 4}, # Tank -> Defender
    "Iori Yagami":  {"syn": ["KOF", "Bruiser", "Mortal Rival"], "cost": 5},
    "Kyo Kusanagi": {"syn": ["KOF", "Mage", "Mortal Rival"],    "cost": 5},

    # --- GLORY LEAGUE ---
    "Aldous":   {"syn": ["Glory League", "Bruiser"],   "cost": 2},
    "Minotaur": {"syn": ["Glory League", "Defender"],  "cost": 3}, # Tank -> Defender
    "Claude":   {"syn": ["Glory League", "Marksman"],  "cost": 4},
    "Granger":  {"syn": ["Glory League", "Assassin"],  "cost": 4},

    # --- METRO ZERO ---
    "Roger": {"syn": ["Metro Zero", "Fighter"], "cost": 2},
    "Ixia":  {"syn": ["Metro Zero", "Marksman", "Stargazer"], "cost": 3}, 
    "Xborg": {"syn": ["Metro Zero", "Dauntless"], "cost": 5}
}

# ------------------------------------------------------
# BAGIAN 2: LIST SPESIAL (LOGIKA GLORY LEAGUE)
# ------------------------------------------------------
# 2 Hero sisa (1 Gold & 5 Gold) diambil acak dari list ini
glory_league_candidates = {
    "1_gold": ["Alucard", "Lolita", "Kula", "Cecilion"],
    "5_gold": ["Akai", "Vexana", "Harley", "Xborg", "Benedetta"]
}

# ------------------------------------------------------
# BAGIAN 3: DATABASE KEKUATAN SINERGI (EFEK & LEVEL)
# ------------------------------------------------------
synergy_db = {
    "Soul Vessels": {
        "description": "Memanggil Dijiang. Dijiang mewarisi % Hybrid ATK.",
        "levels": {
            2: "Dijiang Bintang 1 (5% ATK).",
            4: "Dijiang Bintang 2 (10% ATK).",
            6: "Dijiang Bintang 3 (15% ATK) + Copy Sinergi Terbanyak.",
            10: "Dijiang Bintang 4 (60% ATK) + Copy Sinergi + Skill Chaotic Energy (20k True DMG)."
        }
    },
    "Aspirants": {
        "description": "Peti Kepercayaan. Dapat poin tiap babak. Aktifkan peti untuk hadiah & bonus DMG.",
        "levels": {
            2: "10 Poin/babak. 0.03% DMG/poin.",
            4: "40 Poin/babak. 0.06% DMG/poin.",
            6: "80 Poin/babak. 0.1% DMG/poin.",
            10: "160 Poin/babak. 0.5% DMG/poin. Instan 999 Poin."
        }
    },
    "Shadowcell": {
        "description": "Bonus stat dasar. Tiap Gold refresh nambah stack (max 80). Chance dapat Gold setelah menang.",
        "levels": {
            2: "15% DMG, 15 Def. Stack +0.2% DMG.",
            4: "28% DMG, 30 Def. Stack +0.4% DMG.",
            6: "50% DMG, 60 Def. Stack +0.6% DMG.",
            10: "225% DMG, 250 Def. Stack +2.5% DMG. 100% chance dapat 3 Gold."
        }
    },
    "Starwing": {
        "description": "Petak pesawat antariksa. Hero di petak dapat bonus. Hero Starwing dapat double bonus.",
        "levels": {
            2: "4 Petak. 8% DMG, 300 HP.",
            4: "6 Petak. 12% DMG, 500 HP.",
            6: "Seluruh Board. 16% DMG, 1000 HP.",
            10: "Kapal Induk Bintang. 60% DMG, 6000 HP. Laser 13440 True DMG/8s."
        }
    },
    "Luminexus": {
        "description": "Link damage ke musuh. Musuh terhubung berbagi DMG dan kena debuff Def.",
        "levels": {
            2: "Link 3 musuh. -15 Hybrid Def.",
            4: "Link 6 musuh. -25 Hybrid Def.",
            6: "Link Unlimited. -50 Hybrid Def.",
            10: "Link Unlimited. -600 Hybrid Def. Ledakan 2000 True DMG saat musuh <50% HP."
        }
    },
    "Toy Mischief": {
        "description": "Dapat Equipment Toy Mischief. Bonus Hybrid DEF.",
        "levels": {
            2: "1 Equip. +18 Def.",
            4: "2 Equip. +30 Def.",
            6: "3 Equip. +60 Def.",
            10: "Equip Drastis. +300 Def."
        }
    },
    "Beyond the Clouds": {
        "description": "Skill generate Thunderstone. Seret batu ke hero untuk bonus DMG.",
        "levels": {
            2: "Batu +0.4% DMG.",
            3: "Batu +0.6% DMG."
        }
    },
    "KOF": {
        "description": "Bonus Stat. Kapten (item terbanyak) dapat bonus lebih besar & auto-skill saat teman mati.",
        "levels": {
            2: "5% DMG. Kapten +15% HP.",
            4: "10% DMG. Kapten +25% HP.",
            6: "20% DMG. Kapten +50% HP.",
            11: "200% DMG. Kapten +25000 HP. Aktifkan bonus Duo Mortal Rivals."
        }
    },
    "Mortal Rival": {
        "description": "Bonus saat sendirian. Bonus Duo aktif jika 11 KOF.",
        "levels": {
            1: "Solo: +10% DMG, +5% Dmg Reduc.",
            2: "(Syarat 11 KOF): Duo +100% DMG, +50% Dmg Reduc, +10000 HP."
        }
    },
    "Glory League": {
        "description": "Random 1 Hero (5g) & 1 Hero (1g) join sinergi ini. Bonus Shield & ATK per 3 detik.",
        "levels": {
            2: "+2% Hybrid ATK, 200 Shield.",
            4: "+4% Hybrid ATK, 400 Shield.",
            6: "+10% Hybrid ATK, 900 Shield."
        }
    },
    "Metro Zero": {
        "description": "Bonus stat berdasarkan JUMLAH JENIS sinergi aktif.",
        "levels": {
            2: "+1% ATK per Sinergi. (8 Sinergi: +3% Dmg Reduc). (10 Sinergi: +3% Lifesteal)."
        }
    }
}
# [PASTE KODE DATABASE HERO & SYNERGY DI SINI DULU DI BAGIAN ATAS]
# ... (Hero DB & Synergy DB yang tadi) ...

import streamlit as st
import pandas as pd

# ==========================================
# 4. LOGIKA "OTAK" (ENGINE)
# ==========================================

def get_active_synergies(current_heroes):
    """
    Menghitung sinergi apa saja yang aktif berdasarkan hero di papan.
    """
    counts = {}
    
    # 1. Hitung jumlah tiap sinergi (tag)
    for hero_name in current_heroes:
        data = hero_db.get(hero_name)
        if data:
            for syn in data['synergies']: # Ganti 'syn' ke 'synergies' sesuai DB terakhir
                counts[syn] = counts.get(syn, 0) + 1
    
    # 2. Cek apakah memenuhi syarat aktif
    active_synergies = []
    metro_zero_stack = 0
    
    for syn_name, count in counts.items():
        if syn_name in synergy_db:
            # Cek level aktivasi (misal butuh 2, 4, 6)
            levels = synergy_db[syn_name]['levels']
            # Ambil threshold (angka syarat) dari keys level
            thresholds = sorted(levels.keys())
            
            active_level = 0
            for t in thresholds:
                if count >= t:
                    active_level = t
                else:
                    break
            
            if active_level > 0:
                # Sinergi Aktif!
                desc = levels[active_level]
                active_synergies.append({
                    "name": syn_name,
                    "count": count,
                    "desc": desc
                })
                metro_zero_stack += 1
                
    return active_synergies, metro_zero_stack

def recommend_next_hero(current_heroes):
    """
    Mencari 1 Hero terbaik untuk ditambahkan agar buff Metro Zero maksimal.
    """
    recommendations = []
    current_active, current_stack = get_active_synergies(current_heroes)
    current_active_names = [s['name'] for s in current_active]
    
    # Loop semua hero di database yang belum dipick
    for hero_name in hero_db:
        if hero_name not in current_heroes:
            # Simulasi jika kita tambah hero ini
            simulated_team = current_heroes + [hero_name]
            sim_active, sim_stack = get_active_synergies(simulated_team)
            
            # Hitung kenaikan stack Metro Zero
            stack_gain = sim_stack - current_stack
            
            # Hitung sinergi BARU yang terbuka
            sim_active_names = [s['name'] for s in sim_active]
            new_syns = list(set(sim_active_names) - set(current_active_names))
            
            if stack_gain > 0 or len(new_syns) > 0:
                recommendations.append({
                    "Hero": hero_name,
                    "Cost": hero_db[hero_name]['cost'],
                    "Factions": ", ".join(hero_db[hero_name]['synergies']),
                    "MZ_Gain": stack_gain,
                    "New_Synergies": ", ".join(new_syns)
                })
    
    # Urutkan: Prioritas Gain MZ tertinggi -> Lalu Cost terendah (biar murah)
    recommendations.sort(key=lambda x: (-x['MZ_Gain'], x['Cost']))
    return recommendations

# ==========================================
# 5. TAMPILAN WEB (STREAMLIT UI)
# ==========================================

st.set_page_config(page_title="Metro Zero Solver", layout="wide")

st.title("🤖 Metro Zero Strategy Solver")
st.write("Tools bantu hitung sinergi otomatis biar nggak pusing.")

# --- SIDEBAR: INPUT HERO ---
st.sidebar.header("Tim Kamu Sekarang")
all_heroes = sorted(list(hero_db.keys()))
selected_heroes = st.sidebar.multiselect("Pilih Hero di Papan:", all_heroes)

# --- MAIN AREA: ANALISIS ---
if selected_heroes:
    active_syns, mz_stack = get_active_synergies(selected_heroes)
    
    # 1. TAMPILKAN STATUS METRO ZERO
    st.subheader("📊 Status Metro Zero")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Hero", len(selected_heroes))
    col2.metric("Sinergi Aktif", mz_stack)
    
    # Hitung Buff Metro Zero (Manual Logic sesuai DB)
    mz_atk_buff = mz_stack * 1 # 1% per stack
    mz_dmg_red = "✅ Aktif (3%)" if mz_stack >= 8 else "❌ (Butuh 8)"
    mz_lifesteal = "✅ Aktif (3%)" if mz_stack >= 10 else "❌ (Butuh 10)"
    
    col3.metric("Hybrid ATK Buff", f"+{mz_atk_buff}%")
    st.caption(f"🛡️ Pengurangan DMG: {mz_dmg_red} | ❤️ Lifesteal: {mz_lifesteal}")
    
    # 2. TAMPILKAN DETAIL SINERGI AKTIF
    with st.expander("Lihat Detail Sinergi Aktif"):
        for s in active_syns:
            st.write(f"**{s['name']} ({s['count']})**: {s['desc']}")

    st.divider()

    # 3. REKOMENDASI (THE SOLVER)
    st.subheader("💡 Rekomendasi Hero Selanjutnya")
    st.write("Menambah hero di bawah ini akan membuka sinergi baru/memperkuat Metro Zero:")
    
    recs = recommend_next_hero(selected_heroes)
    
    if recs:
        # Tampilkan sebagai tabel data yang rapi
        df_recs = pd.DataFrame(recs)
        st.dataframe(df_recs, use_container_width=True)
    else:
        st.info("Tidak ada rekomendasi signifikan. Tim kamu mungkin sudah mentok sinerginya.")

else:
    st.warning("👈 Silakan masukkan hero kamu dulu di menu sebelah kiri.")# [PASTE KODE DATABASE HERO & SYNERGY DI SINI DULU DI BAGIAN ATAS]
# ... (Hero DB & Synergy DB yang tadi) ...

import streamlit as st
import pandas as pd

# ==========================================
# 4. LOGIKA "OTAK" (ENGINE)
# ==========================================

def get_active_synergies(current_heroes):
    """
    Menghitung sinergi apa saja yang aktif berdasarkan hero di papan.
    """
    counts = {}
    
    # 1. Hitung jumlah tiap sinergi (tag)
    for hero_name in current_heroes:
        data = hero_db.get(hero_name)
        if data:
            for syn in data['synergies']: # Ganti 'syn' ke 'synergies' sesuai DB terakhir
                counts[syn] = counts.get(syn, 0) + 1
    
    # 2. Cek apakah memenuhi syarat aktif
    active_synergies = []
    metro_zero_stack = 0
    
    for syn_name, count in counts.items():
        if syn_name in synergy_db:
            # Cek level aktivasi (misal butuh 2, 4, 6)
            levels = synergy_db[syn_name]['levels']
            # Ambil threshold (angka syarat) dari keys level
            thresholds = sorted(levels.keys())
            
            active_level = 0
            for t in thresholds:
                if count >= t:
                    active_level = t
                else:
                    break
            
            if active_level > 0:
                # Sinergi Aktif!
                desc = levels[active_level]
                active_synergies.append({
                    "name": syn_name,
                    "count": count,
                    "desc": desc
                })
                metro_zero_stack += 1
                
    return active_synergies, metro_zero_stack

def recommend_next_hero(current_heroes):
    """
    Mencari 1 Hero terbaik untuk ditambahkan agar buff Metro Zero maksimal.
    """
    recommendations = []
    current_active, current_stack = get_active_synergies(current_heroes)
    current_active_names = [s['name'] for s in current_active]
    
    # Loop semua hero di database yang belum dipick
    for hero_name in hero_db:
        if hero_name not in current_heroes:
            # Simulasi jika kita tambah hero ini
            simulated_team = current_heroes + [hero_name]
            sim_active, sim_stack = get_active_synergies(simulated_team)
            
            # Hitung kenaikan stack Metro Zero
            stack_gain = sim_stack - current_stack
            
            # Hitung sinergi BARU yang terbuka
            sim_active_names = [s['name'] for s in sim_active]
            new_syns = list(set(sim_active_names) - set(current_active_names))
            
            if stack_gain > 0 or len(new_syns) > 0:
                recommendations.append({
                    "Hero": hero_name,
                    "Cost": hero_db[hero_name]['cost'],
                    "Factions": ", ".join(hero_db[hero_name]['synergies']),
                    "MZ_Gain": stack_gain,
                    "New_Synergies": ", ".join(new_syns)
                })
    
    # Urutkan: Prioritas Gain MZ tertinggi -> Lalu Cost terendah (biar murah)
    recommendations.sort(key=lambda x: (-x['MZ_Gain'], x['Cost']))
    return recommendations

# ==========================================
# 5. TAMPILAN WEB (STREAMLIT UI)
# ==========================================

st.set_page_config(page_title="Metro Zero Solver", layout="wide")

st.title("🤖 Metro Zero Strategy Solver")
st.write("Tools bantu hitung sinergi otomatis biar nggak pusing.")

# --- SIDEBAR: INPUT HERO ---
st.sidebar.header("Tim Kamu Sekarang")
all_heroes = sorted(list(hero_db.keys()))
selected_heroes = st.sidebar.multiselect("Pilih Hero di Papan:", all_heroes)

# --- MAIN AREA: ANALISIS ---
if selected_heroes:
    active_syns, mz_stack = get_active_synergies(selected_heroes)
    
    # 1. TAMPILKAN STATUS METRO ZERO
    st.subheader("📊 Status Metro Zero")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Hero", len(selected_heroes))
    col2.metric("Sinergi Aktif", mz_stack)
    
    # Hitung Buff Metro Zero (Manual Logic sesuai DB)
    mz_atk_buff = mz_stack * 1 # 1% per stack
    mz_dmg_red = "✅ Aktif (3%)" if mz_stack >= 8 else "❌ (Butuh 8)"
    mz_lifesteal = "✅ Aktif (3%)" if mz_stack >= 10 else "❌ (Butuh 10)"
    
    col3.metric("Hybrid ATK Buff", f"+{mz_atk_buff}%")
    st.caption(f"🛡️ Pengurangan DMG: {mz_dmg_red} | ❤️ Lifesteal: {mz_lifesteal}")
    
    # 2. TAMPILKAN DETAIL SINERGI AKTIF
    with st.expander("Lihat Detail Sinergi Aktif"):
        for s in active_syns:
            st.write(f"**{s['name']} ({s['count']})**: {s['desc']}")

    st.divider()

    # 3. REKOMENDASI (THE SOLVER)
    st.subheader("💡 Rekomendasi Hero Selanjutnya")
    st.write("Menambah hero di bawah ini akan membuka sinergi baru/memperkuat Metro Zero:")
    
    recs = recommend_next_hero(selected_heroes)
    
    if recs:
        # Tampilkan sebagai tabel data yang rapi
        df_recs = pd.DataFrame(recs)
        st.dataframe(df_recs, use_container_width=True)
    else:
        st.info("Tidak ada rekomendasi signifikan. Tim kamu mungkin sudah mentok sinerginya.")

else:
    st.warning("👈 Silakan masukkan hero kamu dulu di menu sebelah kiri.")