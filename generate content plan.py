# -*- coding: utf-8 -*-
"""
Generator Content Plan Official + Personal, 14-22 November 2026.
Alur: shooting -> editing -> QC -> posting dipisah per tanggal (anti-tabrakan).
"""
import math
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from urllib.parse import quote

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT_FILE = "Content_Plan_Official_Personal_14-22_November_2026.xlsx"
YEAR, MONTH = 2026, 11
FONT = "Arial"

HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def d(day):
    return date(YEAR, MONTH, day)


def hari(dt):
    return HARI[dt.weekday()]


# ----------------------------------------------------------------------------
# DATA JADWAL SHOOTING
# ----------------------------------------------------------------------------
SHOOTS = {
    14: ("Trans Park Cibubur", "Eca (tersedia)"),
    16: ("Grand Kamala Lagoon", "Wulan (tersedia)"),
    17: ("Trans Park Juanda", "Belum dikonfirmasi - konfirmasi talent paling lambat 14 Nov"),
    18: ("Spring Lake Bekasi", "Belum dikonfirmasi - konfirmasi talent paling lambat 14 Nov"),
}
BANK_LOC = "Arsip Content Bank (lokasi sesuai arsip - wajib dikonfirmasi)"
BANK_NOTE = "Content Bank (stok sebelum 14 Nov)"

SHEETS = [
    # nama sheet, jenis, niche, warna tab
    ("MV Official", "Official", "Entertainment-first lifestyle & situasi urban anak muda (soft selling)", "1F4E78"),
    ("Loop Property Official", "Official", "Entertainment-first lifestyle & situasi urban anak muda (soft selling)", "1F4E78"),
    ("Samain Official", "Official", "Entertainment-first lifestyle & situasi urban anak muda (soft selling)", "1F4E78"),
    ("Day In Official", "Official", "Entertainment-first lifestyle & situasi urban anak muda (soft selling)", "1F4E78"),
    ("Mia Personal", "Personal", "Personal branding - humor kikuk, polos, awkward", "C55A11"),
    ("Eca Personal", "Personal", "Personal branding - cute, feminin, ekspresif", "C55A11"),
    ("Wulan Personal", "Personal", "Personal branding - komedi spontan, ceplas-ceplos, nuansa Betawi", "C55A11"),
]

PL = "TikTok + Instagram Reels + YouTube Shorts"
PLT = "TikTok (utama) + Instagram Reels"
PLI = "Instagram Reels (utama) + TikTok"

T_BANK = "Sesuai arsip bank (konfirmasi)"
T14 = "Eca + Content Creator"
T16 = "Wulan + Content Creator"
T3 = "Mia + Eca + Wulan (menunggu konfirmasi ketersediaan)"
T_ME = "Mia + Eca (menunggu konfirmasi)"
T_MW = "Mia + Wulan (menunggu konfirmasi)"
T_EW = "Eca + Wulan (menunggu konfirmasi)"
T_MIA = "Mia (menunggu konfirmasi)"

# Pola format umum per arketipe (BUKAN video spesifik)
ARCH = {
    "pov": ("POV skit 1 lokasi",
            "Pola umum POV: teks POV muncul 2 detik pertama, situasi relatable langsung tanpa intro; alur 3 beat (setup, eskalasi, punchline); durasi 15-25 dtk; penutup reaksi diam."),
    "skit": ("Mini skit 3 beat",
             "Pola umum mini skit: dua karakter dengan sifat berlawanan, konflik receh, punchline di 3 detik terakhir; satu sudut kamera agar cepat diedit."),
    "cine": ("Cinematic lifestyle + voice over",
             "Pola umum cinematic: shot close-up benda keseharian, transisi bersih, voice over bernada humor yang mematahkan suasana estetik."),
    "vlog": ("Mini vlog jump-cut",
             "Pola umum mini vlog: potongan cepat per aktivitas, teks jam/aktivitas di layar, musik ringan, ditutup momen lucu."),
    "trend": ("Format tren/reaksi (verifikasi tren aktual sebelum shooting)",
              "Pola umum tren: audio/format yang sedang dipakai banyak kreator diadaptasi ke situasi sendiri; keunggulan ada di kecepatan produksi."),
    "bts": ("Behind the scenes handheld",
            "Pola umum BTS: handheld, spontan, bloopers dan reaksi kru; kesan jujur membangun kedekatan dengan audiens."),
    "story": ("Storytelling singkat + voice over",
              "Pola umum storytelling: hook kalimat pertama berupa pernyataan mengejutkan, 3 adegan, twist ringan di akhir."),
    "review": ("Review/rating spontan to-camera",
               "Pola umum review: to-camera, rating angka di layar, reaksi jujur terhadap benda/makanan kecil; kuat pada karakter pengisi."),
    "react": ("Reaksi/ekspresi close-up",
              "Pola umum reaksi: close-up wajah, perubahan ekspresi bertahap, teks situasi di layar, tanpa dialog panjang."),
}
SEARCH_Q = {
    "pov": "pov situasi kerja lucu", "skit": "skit persahabatan lucu", "cine": "cinematic lifestyle kafe",
    "vlog": "mini vlog keseharian", "trend": "trend reaksi lucu", "bts": "behind the scenes content creator",
    "story": "storytelling pendek relatable", "review": "review jujur jajanan lucu", "react": "reaksi ekspresi lucu",
}

# ----------------------------------------------------------------------------
# DATA IDE
# tuple: (tgl_post, tgl_shoot|None=bank, arketipe, judul, pilar, funnel, konsep,
#         hook, alur, tujuan, talent, adaptasi, script, caption, hashtag, platform)
# ----------------------------------------------------------------------------
IDEAS = {}

IDEAS["MV Official"] = [
    (14, None, "pov", "POV: Nahan Temen Biar Nggak Resign", "Situasi pekerjaan", "TOFU",
     "Dua teman kerja: satu pamit resign, satunya panik menahan dengan alasan receh.",
     "Eh, jangan resign dulu, belum ada yang bisa gantiin kamu.",
     "Teks POV -> teman membawa kardus -> dibujuk dengan alasan receh -> pasrah menunda resign.",
     "Awareness", T_BANK,
     "Setting ruang santai netral; humor dari alasan receh, bukan konflik kerja berat.",
     "A: Aku resign ya.\nB: Hah? Bentar... kamu kan yang tahu password wifi?\nA: Itu alasannya?\nB: Sama cuma kamu yang bisa benerin printer.\nA: Ya udah, minggu depan aja deh.\nB (pelan): Yes.",
     "Belum sempat nyari pengganti, orangnya udah mau resign aja.",
     "#pov #situasikantor #temankerja #relatable #skitkomedi", PL),
    (15, 14, "skit", "Terserah: Debat Pilih Makan Siang", "Persahabatan", "TOFU",
     "Dua sahabat saling melempar 'terserah' soal makan siang sampai semua pilihan ditolak.",
     "Terserah kamu aja. - Beneran terserah?",
     "Tanya mau makan apa -> semua usulan ditolak -> akhirnya mereka memilih menu pertama yang sejak awal diusulkan.",
     "Engagement", T14,
     "Eca memerankan pihak yang menolak semua usulan dengan ekspresi manis tapi tegas.",
     "A: Makan di mana?\nB: Terserah.\nA: Nasi padang?\nB: Jangan, kemarin.\nA: Mie ayam?\nB: Panas.\nA: Ya udah nggak usah makan.\nB: Jangan dong, aku laper!",
     "Tim 'terserah' vs tim 'jangan yang itu'. Kamu masuk tim mana?",
     "#terserah #makansiang #persahabatan #relatable #debatreceh", PL),
    (17, 16, "cine", "Kopi Pertama Sebelum Notifikasi Grup", "Cinematic lifestyle", "MOFU",
     "Pagi sinematik dengan kopi dan cahaya jendela yang dihancurkan oleh notifikasi grup chat.",
     "Pagi itu indah, sampai notifikasi pertama masuk.",
     "Close-up kopi dituang -> cahaya pagi -> ponsel bergetar beruntun -> wajah pasrah ke kamera.",
     "Engagement", T16,
     "Wulan memberi voice over spontan khas Jakarta untuk mematahkan suasana estetik.",
     "VO (Wulan): Pagi tuh tenang... damai... (ponsel getar) Ya Allah, belum juga diminum.\nTeks layar: 47 pesan belum dibaca.",
     "Tenangnya cuma sampai notifikasi pertama.",
     "#cinematic #pagihari #kopipagi #grupchat #lifestyle", PL),
    (18, 17, "vlog", "Mini Vlog: Hari Produktif, Makannya Nanti", "Daily activities", "MOFU",
     "Potongan cepat satu hari sibuk tiga sahabat yang selalu menunda makan.",
     "Jam 7 semangat, jam 3 baru ingat belum makan.",
     "Bangun -> sibuk dengan kegiatan masing-masing -> teks jam di layar -> momen makan terlambat.",
     "Engagement", T3,
     "Porsi tampil seimbang: tiap talent mendapat 1 segmen kegiatan khas karakternya.",
     "Teks: 07.00 Semangat!\n10.00 Masih semangat.\n13.00 Lupa makan.\n15.30 (semua menatap kosong)\nVO: Besok pasti makan tepat waktu.",
     "Hari ini produktif banget, makannya nanti aja.",
     "#minivlog #harikerja #dayinmylife #relatable #lupamakan", PL),
    (20, 17, "trend", "3 Karakter, 1 Situasi: Lampu Mati", "Trend hijacking", "TOFU",
     "Tiga sahabat dengan tiga reaksi berbeda saat lampu tiba-tiba mati.",
     "Lampu mati. Siapa yang panik duluan?",
     "Situasi normal -> lampu mati -> tiga reaksi berurutan -> lampu menyala, semua pura-pura tenang.",
     "Awareness", T3,
     "Setiap karakter tampil sesuai sifat: Mia kikuk, Eca panik manis, Wulan ceplas-ceplos.",
     "Mia: (diam, mematung)\nEca: Aaa! Jangan tinggalin aku!\nWulan: Siapa yang injek kaki gue?!\n(lampu nyala) Semua: Biasa aja kok.",
     "Tiga orang, tiga cara panik. Kamu yang mana?",
     "#tigakarakter #lampumati #sahabat #komedi #reaksilucu", PL),
    (21, 18, "bts", "Take ke-17 Masih Ketawa", "Behind the scenes", "MOFU",
     "Kumpulan gagal take karena tidak bisa menahan tawa.",
     "Ini bukan blooper, ini memang begini kalau syuting bareng.",
     "Take serius -> salah satu tertawa -> kru ikut tertawa -> teks jumlah take.",
     "Engagement", T3,
     "Handheld dari sudut pandang kru; tampilkan momen jujur antar sahabat.",
     "Teks: Take 1... Take 5... Take 12...\nWulan: Udah, kita nyerah, ketawa aja.\nTeks: Take ke-17 baru lolos.",
     "Take ke-17 dan kami masih ketawa.",
     "#behindthescenes #blooper #gagaltake #persahabatan #kreatorkonten", PL),
    (22, 18, "story", "Sore yang Bikin Males Pulang", "Room experience", "BOFU",
     "Cerita pendek tentang sore santai bersama sahabat yang membuat lupa waktu.",
     "Katanya cuma mau mampir sebentar.",
     "Janji mampir sebentar -> obrolan melebar -> langit berubah warna -> semua lupa waktu.",
     "Consideration", T3,
     "Ambil nuansa hangat golden hour; fokus pada kebersamaan, tanpa ajakan atau penyebutan layanan.",
     "VO: Katanya cuma mampir sebentar.\nTeks: 16.00 'Bentar lagi pulang.'\nTeks: 18.30 'Bentar lagi, ya.'\nVO: Besok kerja, tapi sore ini terlalu enak.",
     "Sore ini terlalu enak buat buru-buru pulang.",
     "#soresantai #nongkrong #sahabat #goldenhour #cerita", PL),
]

IDEAS["Loop Property Official"] = [
    (15, 14, "pov", "POV: Teman yang Selalu 'Lima Menit Lagi'", "POV", "TOFU",
     "Satu orang menunggu temannya yang selalu bilang lima menit lagi.",
     "Katanya lima menit lagi. Itu satu jam yang lalu.",
     "Teks jam menunggu -> pesan 'lima menit lagi' -> temannya muncul dengan percaya diri -> reaksi pasrah.",
     "Awareness", T14,
     "Eca sebagai teman yang selalu telat dengan gestur manis tak berdosa.",
     "A (menatap jam): Katanya lima menit.\nB (chat): Lima menit lagi ya.\nA: Itu tadi satu jam lalu.\nB muncul: Maaf, kejebak macet. A: Kamu bahkan belum berangkat.",
     "Lima menit versi dia itu beda sama lima menit versi kita.",
     "#pov #telat #limamenitlagi #relatable #temanhangout", PL),
    (16, 14, "skit", "Grup Chat vs Aslinya", "Relatable", "TOFU",
     "Teman yang super heboh di grup chat tapi sangat pendiam saat bertemu langsung.",
     "Di grup chat dia ratu, aslinya?",
     "Layar chat heboh -> bertemu langsung -> mendadak diam -> pesan chat masuk lagi.",
     "Engagement", T14,
     "Eca memerankan dua versi dirinya dengan kontras ekspresi.",
     "Chat: Gaskeun semua!!! Kumpul jam 7!\nBertemu: (senyum, diam)\nA: Kok diam?\nB: Aku lagi mode offline.\nChat masuk: wkwk asik banget tadi.",
     "Online ngakak, offline mode hemat energi.",
     "#grupchat #introvert #relatable #temanbaru #skit", PL),
    (18, 16, "review", "Wulan Rating Jujur Jajanan Warung", "Lifestyle", "MOFU",
     "Wulan memberi rating angka pada tiga jajanan warung dengan komentar apa adanya.",
     "Rating jujur, nggak ada yang bayar gue.",
     "Perkenalan -> cicip satu-satu -> angka di layar -> peringkat akhir.",
     "Engagement", T16,
     "Pakai jajanan umum dari warung sekitar lokasi, tanpa menyebut brand sponsor.",
     "Wulan: Jajanan pertama... hmm. Enam. Karena kemasannya bagus doang.\nKedua... nah ini baru. Sembilan!\nKetiga... ya ampun, nyesel gue.",
     "Rating jujur: yang paling enak ternyata yang paling murah.",
     "#ratingjujur #jajanan #reviewmakanan #warung #ceplasceplos", PL),
    (19, 17, "story", "Janji Ngopi yang Tertunda", "Storytelling", "MOFU",
     "Dua sahabat yang terus menunda janji ngopi sampai akhirnya benar-benar bertemu.",
     "Kita udah janji ngopi dari bulan kapan ya?",
     "Chat janji -> alasan batal berulang -> akhirnya bertemu -> ngobrol tanpa henti.",
     "Engagement", T_ME,
     "Mia dan Eca sebagai dua sahabat; tonjolkan kehangatan dan sedikit canggung.",
     "VO: Janji ngopi udah tiga bulan.\nTeks: Januari - 'Minggu depan ya.'\nTeks: Maret - 'Aku ada acara.'\nHari ini: 'Akhirnya...' (tertawa)",
     "Janji ngopinya jadi juga, setelah tiga bulan beralasan.",
     "#janjingopi #sahabat #cerita #ngopi #relatable", PL),
    (20, 16, "pov", "POV: Paling Rajin di Grup Tugas", "Situasi pekerjaan", "TOFU",
     "Satu orang mengerjakan semuanya sementara anggota grup lain hanya mengirim stiker.",
     "Tugas kelompok, tapi kerjanya sendirian.",
     "Teks POV -> pembagian tugas -> semua kirim stiker -> orang rajin menatap kamera.",
     "Awareness", T16,
     "Wulan sebagai orang rajin yang ngomel lucu ke kamera.",
     "Wulan: Katanya bagi tugas.\n(ponsel) Anggota 1: Semangat ya! (stiker)\nAnggota 2: Aku ikut doa. (stiker)\nWulan: Doa doang, kerjanya gue.",
     "Yang kerja satu, yang nyemangatin banyak.",
     "#tugaskelompok #pov #situasikerja #relatable #komedi", PL),
    (21, 18, "skit", "Taruhan Bangun Paling Pagi", "Persahabatan", "MOFU",
     "Dua sahabat taruhan siapa bangun lebih pagi, hasilnya sama-sama kesiangan.",
     "Yang kalah traktir es teh.",
     "Taruhan malam hari -> alarm berbunyi -> keduanya terlambat -> saling menyalahkan.",
     "Engagement", T_EW,
     "Eca dan Wulan beradu karakter manis vs ceplas-ceplos.",
     "Eca: Aku pasti bangun jam lima.\nWulan: Gue jam empat!\n(Jam 08.00) Eca: Kamu di mana?\nWulan: Baru bangun. Eca: Aku juga.",
     "Taruhan yang paling nggak ada pemenangnya.",
     "#taruhan #bangunpagi #sahabat #skitkomedi #kesiangan", PL),
    (22, 17, "skit", "Basa-basi Paling Awkward ke Tetangga", "Interaksi sosial", "BOFU",
     "Tiga situasi canggung saat berpapasan dengan tetangga di area bersama.",
     "Senyum, angguk, terus bingung mau ngomong apa.",
     "Berpapasan -> basa-basi kikuk -> ketemu lagi lima menit kemudian -> senyum yang sama.",
     "Consideration", T3,
     "Mia sebagai pusat kecanggungan; Eca dan Wulan merespons dengan gaya masing-masing.",
     "Mia: Eh... udah makan?\nTetangga: Udah.\nMia: Oh... bagus.\n(5 menit kemudian) Mia: Eh, udah makan?\nTetangga: ...Udah.",
     "Basa-basi yang sama dua kali itu skill.",
     "#tetangga #basabasi #awkward #interaksisosial #relatable", PL),
]

IDEAS["Samain Official"] = [
    (14, None, "cine", "Sore, Playlist, dan Segelas Es Teh", "Cinematic lifestyle", "MOFU",
     "Cuplikan sinematik sore santai dengan playlist favorit dan segelas es teh.",
     "Sore ini cukup satu gelas es teh dan playlist yang pas.",
     "Close-up es teh -> jendela sore -> jari mengetuk irama -> senyum kecil.",
     "Engagement", T_BANK,
     "Gunakan footage arsip bertema sore; tambahkan teks ringan sesuai suasana.",
     "VO: Hari ini lumayan melelahkan.\nTeks: Tapi ada es teh.\nVO: Dan playlist yang pas.\nTeks: Cukup.",
     "Kadang cukup es teh dan playlist yang pas.",
     "#esteh #sorehari #playlist #cinematic #slowliving", PL),
    (16, 14, "react", "Ekspresi Pas Dengar 'Rapat Dadakan'", "Trend hijacking", "TOFU",
     "Close-up ekspresi bertahap saat mendengar kabar rapat dadakan.",
     "Pas dengar kata 'rapat dadakan'...",
     "Teks situasi -> wajah tenang -> berubah panik -> pasrah.",
     "Awareness", T14,
     "Eca menampilkan rentang ekspresi manis ke panik tanpa dialog.",
     "Teks: Bos bilang: 'Kita rapat dadakan jam 4.'\n(ekspresi tenang -> kaget -> berpikir -> pasrah)\nTeks: Padahal baru mau pulang.",
     "Wajah ini nggak bisa bohong.",
     "#rapatdadakan #ekspresi #situasikerja #relatable #reaksilucu", PL),
    (17, 16, "skit", "Rebutan Colokan", "Daily activities", "TOFU",
     "Dua orang berebut satu colokan saat baterai ponsel tinggal sedikit.",
     "Baterai 3 persen, colokan cuma satu.",
     "Cari colokan -> berebut -> negosiasi -> kalah oleh pihak ketiga.",
     "Engagement", T16,
     "Wulan sebagai pihak yang paling ngotot; humor dari negosiasi absurd.",
     "A: Aku duluan, tiga persen!\nWulan: Gue dua persen!\nA: Dua persen kan lebih sekarat...\nWulan: Nah itu!\n(Orang ketiga mencolok charger) Keduanya: ...",
     "Perang dunia ketiga: rebutan colokan.",
     "#colokan #baterai #rebutan #relatable #skit", PL),
    (19, 16, "story", "Ngobrol Random dengan Teman Baru di Kafe", "Interaksi sosial", "MOFU",
     "Dua orang asing mengobrol hal receh di meja sebelah dan ternyata punya kenalan yang sama.",
     "Awalnya cuma pinjam sedotan.",
     "Pinjam sedotan -> basa-basi -> kenalan sama -> tertawa.",
     "Engagement", T16,
     "Wulan mudah membuka obrolan dengan gaya Betawi; tonjolkan kehangatan.",
     "Wulan: Mbak, boleh pinjam sedotan?\nOrang: Boleh.\nWulan: Kamu kenal Rina juga?\nOrang: Lho, itu sepupu aku!",
     "Dunia sempit, dan semuanya bermula dari sedotan.",
     "#kafe #temanbaru #ngobrolrandom #sedotan #interaksisosial", PL),
    (20, 18, "skit", "Sahabat yang Selalu Nitip Beli", "Persahabatan", "MOFU",
     "Satu sahabat selalu menitip beli camilan saat temannya keluar.",
     "Aku nggak nitip apa-apa kok. (lalu nitip lima hal)",
     "Pamit keluar -> 'nggak nitip' -> daftar nitip bertambah -> belanjaan penuh.",
     "Engagement", T_EW,
     "Eca dengan gestur memohon manis; Wulan membalas ceplas-ceplos.",
     "Eca: Aku nggak nitip kok.\n(5 detik) Eca: Eh, kalau ada keripik...\nEca: Sama es krim...\nWulan: Itu mah nitip!",
     "Nggak nitip, katanya.",
     "#nitip #camilan #sahabat #persahabatan #skitlucu", PL),
    (21, 17, "story", "Hari Pertama Anak Baru", "Storytelling", "MOFU",
     "Cerita lucu hari pertama anak baru yang bingung dengan peraturan tak tertulis.",
     "Peraturan tak tertulis di hari pertama.",
     "Datang paling pagi -> semua orang santai -> bingung kapan mulai makan siang -> ikut-ikutan santai.",
     "Engagement", T_MIA,
     "Mia sebagai anak baru yang polos; dialog minim, ekspresi dominan.",
     "VO: Hari pertama, datang jam 8 kurang.\nTeks: Semua orang baru datang jam 9.\nVO: Hari kedua, aku ikut telat.\nTeks: Semua orang datang jam 8.",
     "Hari pertama paling nggak ada panduannya.",
     "#haripertama #anakbaru #kantor #storytime #relatable", PL),
    (22, 18, "story", "Kamar Versi Feed vs Versi Asli", "Room experience", "BOFU",
     "Perbandingan lucu kamar yang terlihat rapi di feed dengan kondisi aslinya.",
     "Kamar di feed vs kamar waktu nggak difoto.",
     "Shot rapi -> potong ke kondisi asli -> ekspresi pasrah -> bereskan satu per satu.",
     "Consideration", T_EW,
     "Fokus pada kejujuran humor; tidak menyorot fasilitas atau menawarkan apa pun.",
     "VO: Ini kamarku di feed.\n(potong) VO: Dan ini kamarku besok pagi.\nTeks: Satu-satunya yang konsisten cuma bantalnya.",
     "Yang rapi di feed itu hasil sepuluh menit beres-beres.",
     "#kamar #feedvsreal #kamartidur #relatable #lifestyle", PL),
]

IDEAS["Day In Official"] = [
    (15, 14, "vlog", "Sehari Ala Anak Rantau", "Mini vlog", "MOFU",
     "Mini vlog satu hari anak rantau yang berusaha hemat tapi gagal.",
     "Rencana hari ini: hemat. Realita: jajan.",
     "Bangun pagi -> rencana hemat -> godaan jajan -> saldo dicek -> pasrah.",
     "Engagement", T14,
     "Eca sebagai anak rantau yang polos soal hemat; ekspresi manis saat melihat saldo.",
     "Teks: 08.00 Rencana: bawa bekal.\n10.00 Jajan kopi.\n13.00 Jajan makan siang.\n20.00 Cek saldo. (ekspresi pasrah)",
     "Rencananya hemat, realitanya jajan.",
     "#anakrantau #minivlog #hemat #dayin #relatable", PL),
    (16, 14, "pov", "POV: Alarm Kelima Masih Bohong ke Diri Sendiri", "Humor", "TOFU",
     "Seseorang menunda alarm berkali-kali dengan alasan yang berbeda setiap kali.",
     "Lima menit lagi... (alarm ke-5)",
     "Alarm 1 -> alasan -> alarm 3 -> negosiasi -> alarm 5 -> terlambat.",
     "Awareness", T14,
     "Eca dengan ekspresi mata terpejam dan negosiasi cute dengan alarm.",
     "Alarm 1: Lima menit lagi.\nAlarm 3: Ini hari Sabtu, kan? (bukan)\nAlarm 5: Oke, bangun. (melihat jam) Telat.",
     "Alarm ke-5 itu bukan alarm, itu teman.",
     "#alarm #pov #pagihari #bangunkesiangan #relatable", PL),
    (17, 16, "bts", "Persiapan Take Pertama", "Behind the scenes", "MOFU",
     "Cuplikan persiapan sebelum take pertama: atur kamera, cek properti, bercanda.",
     "Ini yang terjadi sebelum kamera menyala.",
     "Setting kamera -> properti -> latihan dialog -> aba-aba mulai.",
     "Engagement", T16,
     "Wulan menjadi pusat bercandaan kru; handheld spontan.",
     "Wulan: Ini properti kenapa bisa kebalik?\nKru: Kamu yang bawa.\nWulan: Oh iya. (tertawa)\nTeks: Take pertama dimulai 15 menit lebih lambat.",
     "Sebelum action, selalu ada drama kecil.",
     "#behindthescenes #persiapan #kru #syuting #kreator", PL),
    (18, 17, "skit", "Wawancara Kerja Versi Teman Sendiri", "Situasi pekerjaan", "TOFU",
     "Teman memerankan pewawancara kerja yang menanyakan hal absurd.",
     "Kelebihan kamu apa? - Bisa makan banyak.",
     "Salam -> pertanyaan standar -> jawaban absurd -> diterima karena jujur.",
     "Awareness", T_MW,
     "Mia canggung sebagai pelamar; Wulan sebagai pewawancara tanpa filter.",
     "Wulan: Kelebihan kamu apa?\nMia: Eh... bisa makan banyak.\nWulan: Kekurangan?\nMia: Cepat laper.\nWulan: Diterima.",
     "Wawancara paling jujur sepanjang masa.",
     "#wawancarakerja #interview #situasikerja #skit #komedi", PL),
    (19, 18, "cine", "Malam Setelah Lembur", "Cinematic lifestyle", "BOFU",
     "Suasana malam tenang setelah hari panjang dengan cahaya hangat.",
     "Hari panjang, malam akhirnya pelan.",
     "Pintu terbuka -> lampu hangat -> duduk -> napas panjang.",
     "Consideration", T_EW,
     "Cinematic tenang; tanpa dialog, hanya musik dan teks pendek.",
     "Teks: Hari ini panjang.\nTeks: Tapi malam ini pelan.\nVO: Terima kasih untuk diri sendiri.",
     "Hari panjang, malam akhirnya pelan.",
     "#malamhari #selesailembur #cinematic #slowliving #tenang", PL),
    (21, 16, "vlog", "Sabtu: Rencana vs Realita Tiga Sahabat", "Daily activities", "MOFU",
     "Tiga sahabat merencanakan akhir pekan produktif yang berakhir rebahan.",
     "Rencana Sabtu: olahraga, beres-beres, belajar.",
     "Daftar rencana -> aktivitas pertama -> godaan rebahan -> daftar tercoret.",
     "Engagement", T_EW,
     "Eca dan Wulan beradu karakter dalam satu vlog; Mia sebagai pengamat bisu (jika hadir).",
     "Teks: To-do Sabtu: olahraga, beres-beres.\n10.00 Olahraga? (rebahan)\n14.00 Beres-beres? (rebahan)\nTeks: Selesai: rebahan.",
     "To-do list Sabtu: 1 item, rebahan.",
     "#sabtu #akhirpekan #rebahan #minivlog #relatable", PL),
    (22, 17, "story", "Hal Kecil yang Bikin Minggu Sore Berarti", "Storytelling", "BOFU",
     "Cerita pendek tentang kebiasaan kecil di Minggu sore menjelang Senin.",
     "Minggu sore selalu punya rasa yang sama.",
     "Teh hangat -> obrolan ringan -> lipat baju -> menatap langit sore.",
     "Consideration", T3,
     "Pendekatan hangat dan reflektif; tidak menyebut lokasi atau layanan.",
     "VO: Minggu sore itu aneh.\nTeks: Tenang, tapi tahu Senin datang.\nVO: Jadi kita nikmati sorenya.",
     "Minggu sore: tenang sambil pura-pura lupa besok Senin.",
     "#mingguan #soreminggu #storytime #lifestyle #santai", PL),
]

IDEAS["Mia Personal"] = [
    (18, 17, "pov", "Mia Diminta Ngomong di Depan Kamera", "POV", "Tidak Berlaku",
     "Mia mencoba berbicara natural ke kamera tapi terus salah tingkah.",
     "Oke, aku ngomong biasa aja... (kaku)",
     "Mulai -> kaku -> ulang -> tertawa canggung.",
     "Engagement", T_MIA,
     "Humor dari kekakuan Mia yang jujur, tanpa dibuat-buat.",
     "Mia: Halo... aku... eh, halo semua. (menatap kamera lama) Aku lupa mau ngomong apa.",
     "Aku udah latihan di kamar... tapi kameranya beda.",
     "#mia #kikuk #awkward #berbicaradikamera #polos", PLT),
    (19, 17, "skit", "Mia Salah Tingkah Pas Papasan dengan Kenalan", "Interaksi sosial", "Tidak Berlaku",
     "Mia bingung harus menyapa atau pura-pura tidak lihat saat berpapasan dengan kenalan.",
     "Dia lihat aku. Aku lihat dia. Sekarang gimana?",
     "Melihat kenalan -> ragu -> sapaan salah -> kabur ke arah lain.",
     "Engagement", T_MIA,
     "Fokus pada ekspresi kebingungan Mia dengan voice over pikiran.",
     "VO (Mia): Sapa nggak ya?\nMia: Eh... hai.\nKenalan: Hai!\nVO: Terlalu cepat. Terlalu cepat. Jalan lagi.",
     "Aku sapa, tapi kok malah belok ke arah yang salah.",
     "#mia #kikuk #papasan #awkward #polos", PLT),
    (20, 17, "pov", "POV: Pura-pura Sibuk Pas Ketemu Orang yang Dikenal", "Relatable", "Tidak Berlaku",
     "Mia pura-pura menelepon agar tidak perlu ngobrol.",
     "Ponsel bukan buat nelepon, tapi buat kabur.",
     "Melihat kenalan -> ponsel di telinga -> bicara pada ponsel kosong -> ketahuan.",
     "Engagement", T_MIA,
     "Humor kikuk dari akting yang buruk.",
     "Mia (ke ponsel): Halo? Iya... iya... (kenalan lewat) Mia: Aku nggak lagi telepon sebenarnya.",
     "Akting teleponnya nggak lulus.",
     "#mia #polos #pursasibuk #relatable #awkward", PLT),
    (21, 17, "review", "Mia Belajar Pesan Kopi dengan Istilah Keren", "Lifestyle", "Tidak Berlaku",
     "Mia mencoba memesan kopi dengan istilah yang baru dipelajari.",
     "Aku latihan dulu: satu... flat... white.",
     "Latihan di depan cermin -> memesan -> salah sebut -> mendapat minuman yang tidak sesuai.",
     "Engagement", T_MIA,
     "Pakai kafe publik yang tidak ditonjolkan; fokus pada kikuk Mia.",
     "Mia: Satu... flat... wait.\nBarista: Flat white?\nMia: Iya, itu. (lega) Aku latihan dua hari.",
     "Dua hari latihan, tetap salah satu huruf.",
     "#mia #pesankopi #kikuk #lifestyle #polos", PLT),
    (21, 18, "react", "Mia dan Gerakan yang Telat Satu Hitungan", "Trend hijacking", "Tidak Berlaku",
     "Mia mengikuti gerakan tren tapi selalu telat satu hitungan.",
     "Aku udah ngikutin, tapi kenapa beda?",
     "Gerakan benar -> telat -> semakin bingung -> tertawa canggung.",
     "Awareness", T_MIA,
     "Verifikasi tren aktual sebelum shooting; humor dari ketidaksinkronan yang polos.",
     "(Mia mengikuti gerakan)\nTeks: Gerakan 1: aman.\nGerakan 2: telat.\nGerakan 3: arah salah.\nMia: Hehe.",
     "Kakiku sama kepalaku nggak sepakat.",
     "#mia #gerakan #tren #kikuk #awkward", PLT),
    (22, 18, "story", "Mia dan Barang Kecil yang Hilang di Tas Sendiri", "Daily activities", "Tidak Berlaku",
     "Mia mencari kunci yang ternyata ada di tangannya sendiri.",
     "Kunci aku di mana ya? (sambil memegang kunci)",
     "Panik cari -> bongkar tas -> tanya orang -> sadar sedang dipegang.",
     "Engagement", T_MIA,
     "Slapstick ringan tanpa membuat Mia terlihat bodoh; fokus pada kepolosan.",
     "Mia: Kunciku hilang.\nTeman: Itu di tangan kamu.\nMia: (menatap tangan) Oh... ini kunci aku. Makasih.",
     "Kunci sudah ketemu, kepercayaan diri belum.",
     "#mia #kuncihilang #polos #dailylife #relatable", PLT),
    (22, 18, "story", "Pelajaran dari Seminggu Jadi Kreator", "Storytelling", "Tidak Berlaku",
     "Mia menutup minggu dengan merangkum hal yang dipelajari selama pembuatan konten.",
     "Seminggu jadi kreator ternyata nggak gampang.",
     "Kilas balik adegan kikuk -> pelajaran kecil -> tersenyum -> penutup rendah hati.",
     "Engagement", T_MIA,
     "Memberi penutup manis dan menonjolkan kemajuan Mia.",
     "VO (Mia): Aku belajar ngomong ke kamera.\nVO: Aku belajar nggak panik pas salah.\nVO: Dan aku belajar... ternyata aku bisa.",
     "Seminggu ini aku belajar: gugup itu normal.",
     "#mia #refleksi #kreatorkonten #pembelajaran #storytelling", PLT),
]

IDEAS["Eca Personal"] = [
    (16, 14, "vlog", "Morning Routine Cute ala Eca", "Mini vlog", "Tidak Berlaku",
     "Rutinitas pagi Eca yang manis, rapi, dan sedikit dramatis.",
     "Pagi ala Eca dimulai dengan satu hal penting: senyum.",
     "Bangun -> senyum ke cermin -> sarapan kecil -> pilih outfit.",
     "Engagement", T14,
     "Tonjolkan visual, gesture, dan ekspresi wajah Eca.",
     "Teks: 07.00 Bangun!\nVO (Eca): Langkah pertama: senyum.\nTeks: 07.15 Sarapan kecil.\nVO: Hari ini aku akan baik-baik saja.",
     "Pagi ini aku mulai dengan senyum, sisanya nyusul. ✨",
     "#eca #morningroutine #cutelife #minivlog #pagiceria", PLI),
    (17, 14, "review", "Eca Rating Manis: Minuman Pilihan Hari Ini", "Lifestyle", "Tidak Berlaku",
     "Eca memberi rating dengan ekspresi manis pada minuman favoritnya.",
     "Aku kasih bintang, tapi dengan senyum.",
     "Pilih minuman -> cicip -> ekspresi -> bintang di layar.",
     "Engagement", T14,
     "Close-up wajah dan gestur tangan; musik ringan.",
     "Eca: Hmm... manis banget! Aku kasih empat bintang.\nTeks: Kenapa bukan lima? Karena aku masih mau nyari yang lebih enak.",
     "Manisnya pas, kayak aku. 🌸",
     "#eca #ratingminuman #cute #lifestyle #manis", PLI),
    (18, 14, "react", "5 Jenis Senyum Eca", "Daily activities", "Tidak Berlaku",
     "Eca menampilkan lima jenis senyum untuk lima situasi berbeda.",
     "Senyum ke dosen, senyum ke teman, senyum pas lapar...",
     "Teks situasi -> senyum 1 -> senyum 2 -> senyum 3 -> senyum 5.",
     "Engagement", T14,
     "Eca memanfaatkan kekuatan ekspresi wajah; tanpa dialog.",
     "Teks: Senyum ke dosen. (sopan)\nTeks: Senyum ke teman. (lebar)\nTeks: Senyum pas lapar. (lemas)\nTeks: Senyum pas dibilang 'kok kurus'. (bahagia)",
     "Satu wajah, lima senyum. Kamu yang mana? 💕",
     "#eca #senyum #ekspresi #cutegirl #relatable", PLI),
    (19, 14, "pov", "POV: Bilang 'Aku Nggak Lapar' Lalu Ngabisin Makananmu", "Relatable", "Tidak Berlaku",
     "Eca bilang tidak lapar tapi akhirnya ikut mencicipi hingga habis.",
     "Aku nggak lapar kok. (lalu mengambil satu suap)",
     "Menolak -> satu suap -> dua suap -> piring kosong.",
     "Engagement", T14,
     "Ekspresi manis dan berpura-pura polos.",
     "Eca: Aku nggak lapar.\n(1 suap) Eca: Cuma nyicip.\n(5 suap) Eca: Enak sih.\nTeman: Itu piring aku.",
     "Nggak lapar, cuma nyicip... sampai habis. 🍟",
     "#eca #nggaklapar #relatable #makanan #cuteskit", PLI),
    (20, 14, "story", "Outfit Santai Tapi Niat", "Lifestyle", "Tidak Berlaku",
     "Eca memamerkan tiga outfit santai dengan transisi manis.",
     "Santai, tapi tetap niat.",
     "Outfit 1 -> transisi -> outfit 2 -> outfit 3 -> pose penutup.",
     "Engagement", T14,
     "Kekuatan visual dan gestur; musik lembut.",
     "Teks: Outfit hari kampus.\nTeks: Outfit hari ngopi.\nTeks: Outfit hari rebahan tapi tetap cantik.",
     "Santai itu pilihan, tampil manis itu kebiasaan. 🎀",
     "#eca #outfitcheck #ootd #cutestyle #lifestyle", PLI),
    (21, 14, "story", "Hal Kecil yang Bikin Eca Happy", "Storytelling", "Tidak Berlaku",
     "Eca membagikan tiga hal kecil yang membuatnya senang hari itu.",
     "Hal kecil yang bikin hariku lebih baik.",
     "Hal 1 -> hal 2 -> hal 3 -> penutup senyum.",
     "Engagement", T14,
     "Fokus pada kehangatan dan gestur kecil Eca.",
     "VO (Eca): Pertama, es teh yang pas manisnya.\nKedua, teman yang nunggu.\nKetiga, hari ini nggak hujan.",
     "Bahagianya sederhana, tapi cukup banget. 🌷",
     "#eca #bahagiasederhana #storytime #cutelife #syukur", PLI),
    (22, 14, "react", "Tebak Ekspresi Eca", "Interaksi sosial", "Tidak Berlaku",
     "Eca membuat ekspresi dan penonton diminta menebak situasinya.",
     "Tebak, aku lagi ngapain?",
     "Ekspresi 1 -> jeda -> ekspresi 2 -> jawaban muncul.",
     "Engagement", T14,
     "Mengajak interaksi komentar secara wajar lewat tebak-tebakan.",
     "Teks: Ekspresi 1: (mata melebar, bibir manyun)\nTeks: Jawaban: lihat saldo di aplikasi.\nTeks: Ekspresi 2: (senyum lebar) Jawaban: dapat jajan gratis.",
     "Tebak-tebakan dulu: ini ekspresi apa? 🌸",
     "#eca #tebakekspresi #cute #interaktif #ekspresilucu", PLI),
]

IDEAS["Wulan Personal"] = [
    (17, 16, "review", "Satu Take Tanpa Naskah", "Humor", "Tidak Berlaku",
     "Wulan berbicara spontan satu take tanpa naskah tentang topik yang diberikan kru.",
     "Kata kru: ngomong aja, jangan dipikir.",
     "Topik diberikan -> Wulan ngomong -> melenceng -> kembali ke topik.",
     "Engagement", T16,
     "Gaya ceplas-ceplos asli Wulan; edit minimal untuk menjaga spontanitas.",
     "Kru: Topiknya: kopi.\nWulan: Kopi itu... eh, sebelum kopi, lu udah makan belum? Balik ke kopi...",
     "Katanya nggak usah dipikir. Ya udah, gue nggak mikir.",
     "#wulan #spontan #ceplasceplos #satutake #betawi", PLT),
    (18, 16, "review", "Review Jujur Es Kopi Pinggir Jalan", "Lifestyle", "Tidak Berlaku",
     "Wulan mereview es kopi dengan rating jujur dan komentar apa adanya.",
     "Rating jujur, nggak ada yang bayar gue.",
     "Beli -> sedot -> komentar -> rating.",
     "Engagement", T16,
     "Gunakan es kopi umum tanpa brand; fokus pada reaksi Wulan.",
     "Wulan: Hmm... manis. Kemanisan! Tapi enak sih. Tujuh dari sepuluh. Kurang satu karena gue nggak dikasih sedotan lebih.",
     "Tujuh dari sepuluh, sisanya buat sedotan yang pendek.",
     "#wulan #reviewjujur #eskopi #betawi #ceplasceplos", PLT),
    (19, 16, "story", "Kosakata Anak Jakarta untuk Pemula", "Storytelling", "Tidak Berlaku",
     "Wulan mengajarkan kosakata sehari-hari khas Jakarta dengan cara jenaka.",
     "Kamu bukan anak Jakarta kalau belum tahu kata ini.",
     "Kata 1 -> contoh kalimat -> kata 2 -> kata 3 -> kuis kecil.",
     "Engagement", T16,
     "Nuansa Betawi disampaikan dengan sopan dan edukatif, tanpa stereotip negatif.",
     "Wulan: Pelajaran satu: 'kagak' itu 'tidak'.\nPelajaran dua: 'ngapain' itu bukan pertanyaan, itu perasaan.\nPelajaran tiga: 'lu-gue' bebas pakai.",
     "Pelajaran bahasa dulu sebelum lanjut.",
     "#wulan #betawi #bahasajakarta #kosakata #edukasilucu", PLT),
    (20, 16, "pov", "POV: Dikasih Tugas Dadakan Pas Mau Pulang", "Situasi pekerjaan", "Tidak Berlaku",
     "Wulan baru mau pulang tiba-tiba diberi tugas dadakan.",
     "Jam pulang, terus 'eh bisa tolong ini?'",
     "Rapikan tas -> tugas dadakan -> ekspresi -> komentar ceplas-ceplos.",
     "Engagement", T16,
     "Reaksi Wulan yang apa adanya, tetap sopan dan ringan.",
     "Wulan: Udah mau pulang nih.\nAtasan: Tolong ini ya.\nWulan: (diam, lalu senyum) Siap, boss. (pelan) Ya Allah, pulang gue.",
     "Jam pulang itu hak, tapi tugas dadakan itu takdir.",
     "#wulan #tugasdadakan #situasikerja #pov #relatable", PLT),
    (21, 16, "react", "Reaksi Dapat Kabar Traktiran", "Humor", "Tidak Berlaku",
     "Wulan bereaksi berlebihan saat mendengar ada yang mentraktir.",
     "Siapa bilang traktiran? Gue langsung hadir.",
     "Kabar -> reaksi heboh -> bertanya detail -> menyesuaikan pilihan menu.",
     "Awareness", T16,
     "Ekspresi Wulan sangat dominan; durasi pendek untuk jangkauan luas.",
     "Teman: Aku traktir ya.\nWulan: Serius?! Kapan? Di mana? Boleh bawa temen?\nTeman: Cuma es teh.\nWulan: Es tehnya yang gede kan?",
     "Kata 'traktir' bikin gue langsung semangat.",
     "#wulan #traktiran #reaksilucu #darderdor #humor", PLT),
    (21, 16, "story", "Siapa yang Paling Ribet Pas Foto Bareng", "Persahabatan", "Tidak Berlaku",
     "Wulan merating teman-temannya berdasarkan tingkat keribetan saat foto bareng.",
     "Foto bareng itu perang, dan gue wasitnya.",
     "Perkenalan -> teman 1 -> teman 2 -> teman 3 -> peringkat.",
     "Engagement", T16,
     "Gunakan teman nyata dari tim dengan izin; humor ringan tanpa merendahkan.",
     "Wulan: Teman satu: minta ulang lima kali.\nTeman dua: pose lima jenis.\nTeman tiga: sibuk ngecek rambut.\nGue: tinggal senyum doang.",
     "Foto bareng itu butuh kesabaran ekstra.",
     "#wulan #fotobareng #persahabatan #peringkat #komedi", PLT),
    (22, 16, "story", "Cerita Receh: Nyasar Cari Alamat", "Storytelling", "Tidak Berlaku",
     "Wulan menceritakan pengalaman nyasar saat mencari alamat dengan gaya spontan.",
     "Percaya diri banget, ternyata salah jalan.",
     "Percaya diri -> belok salah -> tanya orang -> ternyata berputar.",
     "Engagement", T16,
     "Storytelling spontan dengan potongan b-roll jalanan.",
     "Wulan: Gue udah yakin banget jalannya. Belok kiri. Ternyata kanan. Tanya orang: katanya lurus. Gue ngelewatin tempat yang sama tiga kali.",
     "Aplikasi peta aja bisa kalah sama keyakinan gue.",
     "#wulan #nyasar #storytime #betawi #ceplasceplos", PLT),
]


# ----------------------------------------------------------------------------
# STRUKTUR DATA TERNORMALISASI + PIPELINE
# ----------------------------------------------------------------------------
def build_rows():
    """Ubah IDEAS menjadi list dict dan hitung pipeline produksi."""
    rows_by_sheet = {}
    meta = {s[0]: s for s in SHEETS}
    for sheet, items in IDEAS.items():
        rows = []
        for i, it in enumerate(items, 1):
            (post, shoot, arch, judul, pilar, funnel, konsep, hook, alur, tujuan,
             talent, adaptasi, script, caption, tags, platform) = it
            is_bank = shoot is None
            fmt, pola = ARCH[arch]
            row = {
                "no": i,
                "sheet": sheet,
                "post_date": d(post),
                "shoot_date": None if is_bank else d(shoot),
                "is_bank": is_bank,
                "lokasi": BANK_LOC if is_bank else SHOOTS[shoot][0],
                "akun": sheet,
                "jenis": meta[sheet][1],
                "niche": meta[sheet][2],
                "pilar": pilar,
                "funnel": funnel if meta[sheet][1] == "Official" else "Tidak Berlaku",
                "judul": judul,
                "konsep": konsep,
                "format": fmt,
                "hook": hook,
                "alur": alur,
                "tujuan": tujuan,
                "platform": platform,
                "talent": talent,
                "analisis": pola + " (Pola format umum; belum dicocokkan dengan video spesifik periode 5-12 Nov 2026.)",
                "adaptasi": adaptasi,
                "script": script,
                "caption": caption,
                "hashtag": tags,
                "arch": arch,
            }
            # pipeline
            if is_bank:
                row["sumber"] = BANK_NOTE
                row["edit_due"] = "Siap (arsip)"
                row["quick"] = False
            else:
                gap = (row["post_date"] - row["shoot_date"]).days
                row["quick"] = gap == 1
                if gap == 1:
                    row["sumber"] = "Hasil shooting - quick-turn (edit H+1, posting malam)"
                    row["edit_due"] = row["post_date"]          # final sebelum 16:00 hari posting
                else:
                    row["sumber"] = "Hasil shooting (buffer QC minimal 1 hari)"
                    row["edit_due"] = row["post_date"] - timedelta(days=1)
            rows.append(row)
        rows_by_sheet[sheet] = rows
    assign_post_times(rows_by_sheet)
    add_reviewer_notes(rows_by_sheet)
    return rows_by_sheet


def assign_post_times(rows_by_sheet):
    """Jam posting unik per tanggal agar 7 akun tidak bertabrakan."""
    slots = ["09:00", "11:00", "12:00", "13:30", "15:00", "16:30", "17:30", "18:30",
             "19:00", "19:30", "20:00", "20:30", "21:00", "21:30", "22:00"]
    late = slots[slots.index("19:00"):]
    second = slots[slots.index("17:30"):]
    by_date = defaultdict(list)
    order = [s[0] for s in SHEETS]
    for sheet in order:
        for r in rows_by_sheet[sheet]:
            by_date[r["post_date"]].append(r)
    for dt, items in by_date.items():
        used = set()
        seen_acc = Counter()
        # quick-turn dahulu (slot malam), lalu posting kedua akun yang sama, lalu sisanya
        ordered = sorted(items, key=lambda r: (not r["quick"], r["sheet"]))
        for r in ordered:
            seen_acc[r["sheet"]] += 1
            pool = late if r["quick"] else (second if seen_acc[r["sheet"]] > 1 else slots)
            for s in pool:
                if s not in used:
                    r["jam"] = s
                    used.add(s)
                    break
            else:
                for s in slots:
                    if s not in used:
                        r["jam"] = s
                        used.add(s)
                        break


def add_reviewer_notes(rows_by_sheet):
    for rows in rows_by_sheet.values():
        for r in rows:
            n = ["Referensi video belum terverifikasi: reviewer isi link video asli (publikasi 5-12 Nov 2026) sebelum Approve."]
            if r["is_bank"]:
                n.append("Wajib konfirmasi stok content bank; jika kosong, siapkan b-roll + teks pada 13 Nov atau geser ke 15 Nov.")
            elif r["quick"]:
                n.append("Quick-turn: editor khusus H+1, QC maksimal 16:00, posting malam. Cadangan: geser 1 hari bila gagal QC.")
            if "konfirmasi" in r["talent"].lower() and not r["is_bank"]:
                n.append("Konfirmasi ketersediaan talent paling lambat 14 Nov; bila tidak tersedia, ganti talent sesuai karakter.")
            if "Trend" in r["pilar"] or r["arch"] == "trend":
                n.append("Verifikasi tren/audio aktual sebelum shooting.")
            r["catatan"] = " ".join(n)


# ----------------------------------------------------------------------------
# STYLE
# ----------------------------------------------------------------------------
DARK = "1F2937"
HEADER_A = "0F766E"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Kolom content plan: (header, kunci, lebar)
COLS = [
    ("No.", "no", 6),
    ("Tanggal Shooting", "shoot", 17),
    ("Tanggal Posting", "post", 15),
    ("Hari Posting", "hari", 11),
    ("Deadline Editing Final", "edit_due", 17),
    ("Sumber Materi", "sumber", 26),
    ("Jam Posting", "jam", 10),
    ("Lokasi Shooting", "lokasi", 24),
    ("Nama Akun", "akun", 20),
    ("Jenis Akun", "jenis", 11),
    ("Niche", "niche", 28),
    ("Pilar Konten", "pilar", 18),
    ("TOFU/MOFU/BOFU", "funnel", 14),
    ("Ide/Judul Konten", "judul", 30),
    ("Konsep dan Deskripsi", "konsep", 42),
    ("Format Konten", "format", 24),
    ("Hook", "hook", 32),
    ("Alur Video", "alur", 42),
    ("Tujuan Konten", "tujuan", 15),
    ("Platform", "platform", 22),
    ("Talent", "talent", 26),
    ("Referensi Video Asli", "ref", 34),
    ("Tanggal Publikasi Referensi", "refdate", 20),
    ("Analisis Referensi", "analisis", 50),
    ("Adaptasi untuk Akun", "adaptasi", 42),
    ("Script Lengkap", "script", 55),
    ("Caption / Copywriting", "caption", 40),
    ("Hashtag", "hashtag", 36),
    ("Status", "status", 12),
    ("Catatan Reviewer", "catatan", 50),
]
COL_INDEX = {c[1]: i for i, c in enumerate(COLS, 1)}

# Kolom jadwal shooting -> (judul, kolom awal, kolom akhir) pada grid content plan
SCHED = [
    ("Tanggal", "B", "B"),
    ("Hari", "C", "C"),
    ("Lokasi", "D", "E"),
    ("Talent", "F", "F"),
    ("Akun yang Diproduksi", "G", "H"),
    ("Target Jumlah Video", "I", "I"),
    ("Jenis Konten", "J", "L"),
    ("Catatan Produksi", "M", "O"),
]
SCHED_HEADERS = [s[0] for s in SCHED]

ROW_SHEET_TITLE = 1
ROW_SCHED_TITLE = 2
ROW_SCHED_HEADER = 3
ROW_SCHED_FIRST = 4
ROW_SCHED_NOTE = 8
ROW_PLAN_TITLE = 10
ROW_PLAN_HEADER = 11
ROW_PLAN_FIRST = 12


def create_workbook():
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def create_worksheet(wb, name, tab_color):
    ws = wb.create_sheet(title=name)
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = False
    return ws


def font(bold=False, color="000000", size=10, underline=None):
    return Font(name=FONT, bold=bold, color=color, size=size, underline=underline)


def apply_header_style(ws, row, c1, c2, fill_hex=DARK):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=fill_hex)
        cell.font = font(True, WHITE, 10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def apply_table_borders(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = BORDER


def col_idx(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def est_lines(text, width):
    if text is None:
        return 1
    total = 0
    for part in str(text).split("\n"):
        total += max(1, math.ceil(len(part) / max(1, width * 1.05)))
    return total


def add_shooting_schedule(ws, sheet_name, rows):
    ws.merge_cells(start_row=ROW_SHEET_TITLE, start_column=1, end_row=ROW_SHEET_TITLE, end_column=15)
    t = ws.cell(row=ROW_SHEET_TITLE, column=1,
                value=f"{sheet_name.upper()} - CONTENT PLAN 14-22 NOVEMBER 2026")
    t.font = font(True, WHITE, 14)
    t.fill = PatternFill("solid", fgColor=DARK)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[ROW_SHEET_TITLE].height = 28

    ws.merge_cells(start_row=ROW_SCHED_TITLE, start_column=1, end_row=ROW_SCHED_TITLE, end_column=15)
    s = ws.cell(row=ROW_SCHED_TITLE, column=1, value="BAGIAN A: JADWAL SHOOTING")
    s.font = font(True, HEADER_A, 11)
    s.alignment = Alignment(vertical="center")
    ws.row_dimensions[ROW_SCHED_TITLE].height = 22

    for title, a, b in SCHED:
        ca, cb = col_idx(a), col_idx(b)
        if ca != cb:
            ws.merge_cells(start_row=ROW_SCHED_HEADER, start_column=ca, end_row=ROW_SCHED_HEADER, end_column=cb)
        ws.cell(row=ROW_SCHED_HEADER, column=ca, value=title)
        apply_header_style(ws, ROW_SCHED_HEADER, ca, cb, HEADER_A)
    ws.row_dimensions[ROW_SCHED_HEADER].height = 26

    for i, day in enumerate(sorted(SHOOTS)):
        r = ROW_SCHED_FIRST + i
        lokasi, talent_avail = SHOOTS[day]
        mine = [x for x in rows if x["shoot_date"] == d(day)]
        formats = sorted({x["format"].split(" (")[0] for x in mine})
        if mine:
            catatan = (f"Edit mulai {d(day).day + 1} Nov; posting dari hasil shoot ini "
                       f"{', '.join(str(x['post_date'].day) for x in sorted(mine, key=lambda z: z['post_date']))} Nov.")
            akun = sheet_name
            jenis = ", ".join(formats)
        else:
            catatan = "Tidak ada produksi untuk akun ini pada sesi ini; sesi dipakai akun lain."
            akun = f"{sheet_name} (tidak ada produksi)"
            jenis = "-"
        values = [d(day), hari(d(day)), lokasi, talent_avail, akun, len(mine), jenis, catatan]
        for (title, a, b), v in zip(SCHED, values):
            ca, cb = col_idx(a), col_idx(b)
            if ca != cb:
                ws.merge_cells(start_row=r, start_column=ca, end_row=r, end_column=cb)
            cell = ws.cell(row=r, column=ca, value=v)
            cell.font = font()
            cell.alignment = Alignment(horizontal="center" if title in ("Tanggal", "Hari", "Target Jumlah Video") else "left",
                                       vertical="top", wrap_text=True)
            if isinstance(v, date):
                cell.number_format = "DD MMM YYYY"
        apply_table_borders(ws, r, r, 2, 15)
        ws.row_dimensions[r].height = 42

    ws.merge_cells(start_row=ROW_SCHED_NOTE, start_column=2, end_row=ROW_SCHED_NOTE, end_column=15)
    n = ws.cell(row=ROW_SCHED_NOTE, column=2,
                value=("Aturan pipeline: shooting hanya 14, 16, 17, 18 Nov. Editing mulai H+1, QC minimal H-1 sebelum posting "
                       "(kecuali quick-turn: QC maksimal 16:00 hari posting, posting malam). Posting berjalan setiap hari 14-22 Nov; "
                       "14 Nov memakai content bank. Jam posting dibuat unik per tanggal agar 7 akun tidak bertabrakan."))
    n.font = font(False, "404040", 9)
    n.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[ROW_SCHED_NOTE].height = 40


def ref_search_url(arch):
    return "https://www.tiktok.com/search?q=" + quote(SEARCH_Q[arch])


def add_content_plan(ws, rows):
    ncol = len(COLS)
    ws.merge_cells(start_row=ROW_PLAN_TITLE, start_column=1, end_row=ROW_PLAN_TITLE, end_column=15)
    p = ws.cell(row=ROW_PLAN_TITLE, column=1, value="BAGIAN B: CONTENT PLAN")
    p.font = font(True, DARK, 11)
    ws.row_dimensions[ROW_PLAN_TITLE].height = 22

    for i, (h, key, w) in enumerate(COLS, 1):
        ws.cell(row=ROW_PLAN_HEADER, column=i, value=h)
    apply_header_style(ws, ROW_PLAN_HEADER, 1, ncol)
    ws.row_dimensions[ROW_PLAN_HEADER].height = 34

    for idx, r in enumerate(sorted(rows, key=lambda z: z["no"])):
        rr = ROW_PLAN_FIRST + idx
        vals = dict(r)
        vals["shoot"] = r["shoot_date"] if r["shoot_date"] else "Content Bank"
        vals["post"] = r["post_date"]
        vals["hari"] = hari(r["post_date"])
        vals["ref"] = "Belum terverifikasi - tautan pencarian (BUKAN video referensi), isi link video asli"
        vals["refdate"] = "Tidak terverifikasi"
        vals["status"] = "Draft"
        for c, (h, key, w) in enumerate(COLS, 1):
            cell = ws.cell(row=rr, column=c, value=vals.get(key))
            cell.font = font()
            cell.alignment = Alignment(wrap_text=True, vertical="top",
                                       horizontal="center" if key in ("no", "shoot", "post", "hari", "edit_due", "jam", "jenis", "funnel", "status") else "left")
            if isinstance(vals.get(key), date):
                cell.number_format = "DD MMM YYYY"
        # tinggi baris
        lines = max(est_lines(vals.get(k), w) for (_, k, w) in COLS)
        ws.row_dimensions[rr].height = min(400, max(60, lines * 13.5))
    last = ROW_PLAN_FIRST + len(rows) - 1
    apply_table_borders(ws, ROW_PLAN_FIRST, last, 1, ncol)
    return last


def apply_hyperlinks(ws, rows, last):
    c = COL_INDEX["ref"]
    for idx, r in enumerate(sorted(rows, key=lambda z: z["no"])):
        cell = ws.cell(row=ROW_PLAN_FIRST + idx, column=c)
        cell.hyperlink = ref_search_url(r["arch"])
        cell.font = font(False, "0563C1", 10, "single")


def apply_column_widths(ws):
    for i, (_, _, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def apply_status_dropdown(ws, last):
    col = get_column_letter(COL_INDEX["status"])
    dv = DataValidation(type="list", formula1='"Draft,Approve,Revisi,Cancel"', allow_blank=False,
                        showErrorMessage=True, errorTitle="Status tidak valid",
                        error="Pilih salah satu: Draft, Approve, Revisi, Cancel.")
    ws.add_data_validation(dv)
    dv.add(f"{col}{ROW_PLAN_FIRST}:{col}{last}")


def apply_conditional_formatting(ws, last):
    col = get_column_letter(COL_INDEX["status"])
    rng = f"{col}{ROW_PLAN_FIRST}:{col}{last}"
    for text, color in (("Draft", "FFF2CC"), ("Approve", "C6EFCE"), ("Revisi", "F8CBAD"), ("Cancel", "FFC7CE")):
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{text}"'],
                            fill=PatternFill("solid", start_color=color, end_color=color, bgColor=color)))


def apply_freeze_panes(ws):
    ws.freeze_panes = ws.cell(row=ROW_PLAN_FIRST, column=2)


def apply_autofilter(ws, last):
    ws.auto_filter.ref = f"A{ROW_PLAN_HEADER}:{get_column_letter(len(COLS))}{last}"


def apply_print_settings(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def save_workbook(wb, path=OUTPUT_FILE):
    wb.save(path)
    return path


# ----------------------------------------------------------------------------
# VALIDASI
# ----------------------------------------------------------------------------
FORBIDDEN = re.compile(r"transit|sewa|per jam|\bharian\b|jangka pendek|bulanan", re.I)
PROMO = re.compile(r"apartemen|apartment|promo|diskon|sewa|booking|kunjungi|harga|cicil|kpr|dp 0|hubungi|dm kami|unit ", re.I)


def validate_workbook(path, rows_by_sheet):
    report = []
    ok_all = True

    def check(name, cond, detail=""):
        nonlocal ok_all
        ok_all &= bool(cond)
        report.append((name, "LULUS" if cond else "GAGAL", detail))

    wb = load_workbook(path)
    names = [s[0] for s in SHEETS]
    check("File dapat dibuka kembali", True, path)
    check("Jumlah worksheet tepat 7", len(wb.sheetnames) == 7, str(len(wb.sheetnames)))
    check("Nama worksheet sesuai", wb.sheetnames == names, ", ".join(wb.sheetnames))

    headers = [c[0] for c in COLS]
    all_titles, all_concepts = [], []
    posts_per_date = Counter()
    times_per_date = defaultdict(list)
    loc_ok = date_ok = day_ok = pipe_ok = True
    hyper_cnt = dv_ok = cf_ok = True
    hyper_total = 0
    forbid_hit, promo_hit = [], []
    right_most = True
    counts = {}
    for name in names:
        ws = wb[name]
        sched_ok = ws.cell(row=ROW_SCHED_TITLE, column=1).value == "BAGIAN A: JADWAL SHOOTING"
        plan_ok = ws.cell(row=ROW_PLAN_TITLE, column=1).value == "BAGIAN B: CONTENT PLAN"
        check(f"[{name}] Jadwal shooting & content plan ada", sched_ok and plan_ok)
        hdr = [ws.cell(row=ROW_PLAN_HEADER, column=i).value for i in range(1, len(COLS) + 1)]
        check(f"[{name}] Semua kolom wajib tersedia", hdr == headers)
        right_most &= hdr[-1] == "Catatan Reviewer"
        # jadwal shooting
        for i, day in enumerate(sorted(SHOOTS)):
            r = ROW_SCHED_FIRST + i
            v_date = ws.cell(row=r, column=2).value
            v_hari = ws.cell(row=r, column=3).value
            v_loc = ws.cell(row=r, column=4).value
            if not (v_date.date() == d(day) and v_hari == hari(d(day)) and v_loc == SHOOTS[day][0]):
                loc_ok = False
        # baris content plan
        n = 0
        r = ROW_PLAN_FIRST
        while ws.cell(row=r, column=1).value is not None:
            row = {h: ws.cell(row=r, column=i).value for i, h in enumerate(headers, 1)}
            n += 1
            post = row["Tanggal Posting"]
            if not (isinstance(post, datetime) and date(2026, 11, 14) <= post.date() <= date(2026, 11, 22)):
                date_ok = False
            else:
                if row["Hari Posting"] != hari(post.date()):
                    day_ok = False
                posts_per_date[post.date()] += 1
                times_per_date[post.date()].append(row["Jam Posting"])
            sh = row["Tanggal Shooting"]
            if isinstance(sh, datetime):
                if sh.date() not in [d(x) for x in SHOOTS] or row["Lokasi Shooting"] != SHOOTS[sh.day][0]:
                    loc_ok = False
                if (post - sh).days < 1:
                    pipe_ok = False
            else:
                if sh != "Content Bank" or post.date() > d(15):
                    pipe_ok = False
            text_all = " ".join(str(row[k]) for k in headers if k not in ("Catatan Reviewer", "Referensi Video Asli"))
            if FORBIDDEN.search(text_all) or FORBIDDEN.search(str(row["Catatan Reviewer"])):
                forbid_hit.append((name, r))
            cap = f"{row['Caption / Copywriting']} {row['Hashtag']}"
            if PROMO.search(cap):
                promo_hit.append((name, r))
            if ws.cell(row=r, column=COL_INDEX["ref"]).hyperlink is not None:
                hyper_total += 1
            all_titles.append(re.sub(r"\W+", " ", row["Ide/Judul Konten"].lower()).strip())
            all_concepts.append(row["Konsep dan Deskripsi"].strip().lower())
            if row["Status"] != "Draft":
                dv_ok = False
            r += 1
        counts[name] = n
        last = ROW_PLAN_FIRST + n - 1
        dvs = [dv for dv in ws.data_validations.dataValidation if dv.type == "list"]
        dv_ok &= bool(dvs) and "Draft,Approve,Revisi,Cancel" in dvs[0].formula1
        cf_ok &= len(ws.conditional_formatting) > 0
        check(f"[{name}] Jumlah ide 5-7", 5 <= n <= 7, str(n))
        check(f"[{name}] Freeze panes & autofilter", ws.freeze_panes is not None and ws.auto_filter.ref is not None,
              f"{ws.freeze_panes} | {ws.auto_filter.ref}")

    check("Tanggal posting dalam 14-22 Nov 2026", date_ok)
    check("Hari sesuai tanggal", day_ok)
    check("Tanggal shooting sesuai lokasi (4 sesi)", loc_ok)
    check("Posting H+1 atau lebih setelah shooting; bank hanya 14-15 Nov", pipe_ok)
    all_days = [d(x) for x in range(14, 23)]
    missing = [x for x in all_days if posts_per_date[x] == 0]
    check("Setiap tanggal 14-22 Nov memiliki posting", not missing,
          ", ".join(f"{x.day}:{posts_per_date[x]}" for x in all_days))
    dup_time = [(dt, t) for dt, ts in times_per_date.items() for t, c in Counter(ts).items() if c > 1]
    check("Jam posting unik per tanggal (tanpa tabrakan)", not dup_time, str(dup_time))
    check("Status dropdown (DataValidation) tersedia", dv_ok)
    check("Conditional formatting tersedia", cf_ok)
    check("Hyperlink tersimpan di kolom referensi", hyper_total == sum(counts.values()), f"{hyper_total} sel (tautan pencarian)")
    check("Catatan Reviewer paling kanan", right_most)
    check("Tidak ada judul duplikat", len(all_titles) == len(set(all_titles)))
    check("Tidak ada konsep duplikat", len(all_concepts) == len(set(all_concepts)))
    check("Tidak ada caption/hashtag promosi", not promo_hit, str(promo_hit))
    check("Tidak ada skema penyewaan terlarang", not forbid_hit, str(forbid_hit))
    off = sum(v for k, v in counts.items() if "Official" in k)
    per = sum(v for k, v in counts.items() if "Personal" in k)
    check("Minimal 28 ide official, 21 ide personal, 49 total", off >= 28 and per >= 21 and off + per >= 49,
          f"official={off}, personal={per}, total={off + per}")
    return ok_all, report, counts, posts_per_date


def main():
    rows_by_sheet = build_rows()
    wb = create_workbook()
    for name, jenis, niche, color in SHEETS:
        ws = create_worksheet(wb, name, color)
        rows = rows_by_sheet[name]
        add_shooting_schedule(ws, name, rows)
        last = add_content_plan(ws, rows)
        apply_column_widths(ws)
        apply_hyperlinks(ws, rows, last)
        apply_status_dropdown(ws, last)
        apply_conditional_formatting(ws, last)
        apply_freeze_panes(ws)
        apply_autofilter(ws, last)
        apply_print_settings(ws)
    save_workbook(wb)
    ok, report, counts, posts = validate_workbook(OUTPUT_FILE, rows_by_sheet)

    print("=== RINGKASAN IDE PER WORKSHEET ===")
    for k, v in counts.items():
        print(f"{k}: {v}")
    print("=== POSTING PER TANGGAL ===")
    for x in range(14, 23):
        print(f"{x} Nov: {posts[d(x)]}")
    quick = sum(1 for rs in rows_by_sheet.values() for r in rs if r["quick"])
    bank = sum(1 for rs in rows_by_sheet.values() for r in rs if r["is_bank"])
    load = Counter(r["shoot_date"].day for rs in rows_by_sheet.values() for r in rs if r["shoot_date"])
    print(f"Quick-turn: {quick} | Content bank: {bank} | Beban shooting: {dict(load)}")
    print("=== HASIL VALIDASI ===")
    for name, status, detail in report:
        print(f"[{status}] {name} {('- ' + detail) if detail else ''}")
    print("SEMUA LULUS" if ok else "ADA YANG GAGAL")
    return ok


if __name__ == "__main__":
    main()
