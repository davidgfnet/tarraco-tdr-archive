
Tàrraco: Creació d'un joc d'ordinador
=====================================

Aquest repo ve d'un backup que tenia per casa (data del 12/05/2008) però els
continguts són probablement molt més antics (mínim 1 any).

L'objectiu de pujar-ho a github és mantenir una còpia en algun lloc on no es
perdi i que, potser, sigui d'utilitat algú. Com a mínim des d'un punt de
vista "històric" si més no.

Per als que no tenen context: Tàrraco és un joc inacabat que vaig desenvolupar
quan tenia 15-16 anys a l'escola com a treball de recerca de batxillerat.


Continguts
----------

He separat en directories les diferents parts del backup que, per motius
"històrics" estava originalment barrejat:

 - Doc: Arxius del treball mateix. Realitzats originalment amb InDesign
   d'Adobe, i exportats en PDF per a entrega i impressió.
   S'inclou també el video/flash de la presentació oral.
 - Promo: Alguns materials com ara screenshots i vídeos que vaig fer abans,
   durant i després del desenvolupament del joc.
 - CD-Release-Loader: Petit projecte de CD-Launcher (autorun.exe) pel CD.
 - 3d-models: els models del joc en format 3DS Max, el més important dels
   quals és el "tarraco", aixi com d'altres ("soldado"). Molts objectes
   són instanciats pel joc i, per tant, s'exporten a banda. També hi ha
   les textures i una barrija-barreja d'altres coses.
 - src: Conté el projecte VB6 original, codename "gameproject" que conté
   el codi font, així com els packs de "resources" que contenen textures
   i models empaquetats. A dins trobareu els sub-projectes:
    - Coldet: Llibreria C++ de col.lisions Open Source.
    - Cal3D: Llibreria C++ de animació esqueletal, també OSS.
    - engine: Llibreria C++ que fa ús de Coldet i és cridada pel codi VB
      per a processar tota la feina de col.lisió del joc. És en C++ per
      tal de ser més ràpida així com de poder usar Coldet.
    - cal3dvb: Adaptador C++/VB6 de la lliberia Cal3D. Bàsicament es un
      proveïdor de COM objects (si no recordo malament) que permet al codi
      VB6 crear objectes Cal3D i fer certes crides per animar-los així com
      obtenir els vèrtexs i els triangles associats a cada moment (per render).

Tots aquests projectes tenen les seves pròpies llicències. Per al codi que vaig
escriure jo, no hi havia cap llicència en el seu moment, pel que ara aprofito
per assignar-li una llicència de codi lliure GPL o, si no és possible (en el
cas de modifications d'altres llibreries no-GPL) en la llicència que
s'escaigui (MIT, BSD, etc...). Per qualsevol dubte millor preguntar :)

La llicència de les textures és un poc desconeguda, perquè n'hi ha de totes
bandes. La majoria però venen de davegh.com (de David Gurrea).


