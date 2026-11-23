# Telegram Bakery Bot (demo)

Bot Telegram pentru o cofetărie, care folosește un model AI local (Ollama).

## Ce face
- Răspunde la întrebări despre produse, prețuri și program
- Preia comenzi pas cu pas: nume, telefon (validat), livrare, confirmare
- Calculează totalul în cod, nu cu AI-ul
- Salvează comenzile în Excel și îl notifică pe patron pe Telegram

## Cum funcționează
Modelul AI răspunde la întrebări și recunoaște produsele din mesaj.
Codul conduce comanda și validează datele, deci botul nu poate
confirma o comandă incompletă sau inventa produse.

## Instalare
pip install pyTelegramBotAPI ollama openpyxl
ollama pull llama3.2

## Rulare (PowerShell)
$env:TELEGRAM_TOKEN="tokenul_de_la_BotFather"
$env:PATRON_CHAT_ID="id_ul_tau_numeric"
python bot.py

## Notă
Datele din catalog.txt sunt fictive, pentru demo.
