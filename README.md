# CONVERTOR XLSX U XML


#### PROGRAM OMOGUĆAVA PRETVARANJE XLSX FAJLA U XML FAJL KOJI JE MOGUĆE UČITATATI U PSIT MPP2

### PREDUSLOVI
1. Instaliran 64-bitni Python programski jezik, link za preuzimanje: https://www.python.org/downloads/windows/
2. Instaliran 64-bitni ODBC drajver. Link za preuzimanje: https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver17




### PODEŠAVANJE KONEKCIJE (slika 1)
Konekciju ka SQL Serveru je neophodno podesiti kako bi program preuzeo podatke o kontnom planu, kao i preduzećima.
Način podešavanja je sledeći
1. Otvoriti program za administraciju: MPP 2.4 Administracija
2. U delu RAD SA BAZOM/PARAMETRI POVEZIVANJA, nalazi se naziv servera
3. U delu PORTOVI/PORT NA KOME RADI SERVER nalazi se broj porta koji je neophodno uneti u program
4. Kliknuti na dugme TEST KONEKCIJE, program će prikazati poruku o izvršenoj konekciji (slika 2)

### RAD SA PROGRAMOM
1. Spremti XLSX fajl u kojem su evidentirane poslovne promene. Nacrt Excel fajla mora biti preuzet sa ovog sajta  -  template_knjizenje_sa_kontom.xlsx (slika 3)
2. U delu programa ULAZNI XLSX, učitati xlsx fajl (slika 4)
3. Ukoliko je program ispravno učitao xlsx fajl, isit će biti prikazan u delu PREGLED XLSX (prvih 20 redova) (slika 5)
4. U delu IZLAZ I GENERISANJE, kliknuti na dugme SAČUVAJ KAO, sačuvati XML, a zatim klknuti na dugme GENERIŠI XML (slika 6)

### UVOZ XML U MPP
1. Generiran XML fajl se učitava u delu programa UVOZ I IZVOZ/DOKUMENTI/UVOZ DOKUMENATA (slika 7)
2. U čarobnjaku kliknuti na dugme SLEDEĆE (slika 8)
3. U delu IZBOR ŽELJENOG UVOZA, izabrati NALOG ZA KNJIŽENJE (slika 9)
4. Izabrati putanju do XML fajla (slika 10)
5. Ukoliko prilikom uvoza MPP prijavi da se šifre preduzeća ne poklapaju, nastaviti sa uvozom
6. Selektovati fajl (slika 11), i kliknuti na sledeće dok MPP ne učita fajl
7. Proveriti učitan nalog za knjiženje


### 

<img width="1395" height="1027" alt="image" src="https://github.com/user-attachments/assets/aeb34677-b81f-4dc4-977f-523f3e54e9e8" />

<img width="594" height="271" alt="image" src="https://github.com/user-attachments/assets/2395b877-00ca-42db-a4c6-7d77a05d3000" />

<img width="1054" height="628" alt="image" src="https://github.com/user-attachments/assets/1206d78e-d3cb-4c48-a0e3-084db5120012" />

<img width="468" height="304" alt="image" src="https://github.com/user-attachments/assets/f9b05a32-511e-4697-b9a2-89c31b389461" />

<img width="1378" height="289" alt="image" src="https://github.com/user-attachments/assets/e61da020-1c49-43c4-b0ac-22686caa62a2" />

<img width="1370" height="216" alt="image" src="https://github.com/user-attachments/assets/f6b8345b-9113-48f4-b9c4-5342e95e4848" />

<img width="524" height="276" alt="image" src="https://github.com/user-attachments/assets/da71f6de-5439-4df7-bec0-eb477bcb0054" />

<img width="741" height="555" alt="image" src="https://github.com/user-attachments/assets/e67e061f-bddb-42c1-aeef-f82be4934f72" />

<img width="792" height="564" alt="image" src="https://github.com/user-attachments/assets/6dbe008b-ebd7-45da-af21-b3f43bfb46bd" />

<img width="730" height="549" alt="image" src="https://github.com/user-attachments/assets/8ea642d3-bdd2-4ca4-9495-a02e074fa3c4" />

<img width="1257" height="308" alt="image" src="https://github.com/user-attachments/assets/b5f45721-e884-494f-8eaa-d42585b4bb9a" />

<img width="985" height="739" alt="image" src="https://github.com/user-attachments/assets/02235aef-a507-4e0e-821d-7165e0ed98ad" />




