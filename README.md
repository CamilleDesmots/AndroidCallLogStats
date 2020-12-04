# AndroidCallLogStats

Petite application en ligne de commande qui à partir d'un répertoire donné va 
rechercher tout les fichier avec l'extension ".clbu"

Un fichier ".clbu" est un fichier d'historique Androïd exporté au format CSV.

Exemple d'extrait de contenu de fichier ".clbu" :

  +33675700011;2;1557077177394;0;0;1;

La 1ère valeur "+33675700011" est le numéro de téléphone appelé ou appelant.
La 2ème valeur "2" signifie qu'il s'agit d'un appel sortant. Pour un appel entrant c'est "1" et pour un appel en absence c'est "3".
La 3ème valeur correspond à la date et l'heure précise en millisecondes au format Epoch.
La 4ème valeur correspond à la durée de l'appel en secondes.


# Utilisation

Télécharger le fichier [Fichier JAR à télécharger](./download/AndroidCallLogStats-1.0-SNAPSHOT-jar-with-dependencies.jar) dans un répertoire

Ensuite lancer la commande suivante : 

    java -jar -DclbuFolder="/Users/camilledesmots/Google Drive/oed.lpo35/Historique des appels/Log des appels/" -DresultFolder="./" AndroidCallLogStats-1.0-SNAPSHOT-jar-with-dependencies.jar

Le paramètre **"clbuFolder="** correspond au répertoire ou se trouvent les fichiers ".clbu" à analyser.
Le paramètre **"resultFolder="** correspond au répertoire ou va être généré le fichier EXCEL.

Exemple de nom de fichier EXCEL généré : "Statistiques des appels au 2020-12-01 généré le 2020-12-03_23-34-26.xlsx".
La date "2020-12-01" correspond à la date de l'appel le plus récent trouvé dans les fichiers.
 