# AndroidCallLogStats

Petite application JAVA en ligne de commande qui à partir d'un répertoire donné va 
rechercher tous les fichier avec l'extension ".clbu". Elle va lire chacun de ces fichiers pour produire des statistiques sous la forme d'un fichier au format Microsoft Excel.

Un fichier ".clbu" est un fichier d'historique Androïd exporté au format CSV.

Exemple d'extrait de contenu de fichier ".clbu" :

  +33675700011;2;1557077177394;0;0;1;

 - La 1ère valeur "+33675700011" est le numéro de téléphone appelé ou appelant.
 - La 2ème valeur "2" signifie qu'il s'agit d'un appel sortant. Pour un appel entrant c'est "1" et pour un appel en absence c'est "3".
 - La 3ème valeur correspond à la date et l'heure précise en millisecondes au format Epoch.
 - La 4ème valeur correspond à la durée de l'appel en secondes.

# Utilisation

Télécharger le fichier [Fichier JAR à télécharger](./download/AndroidCallLogStats-1.0-SNAPSHOT-jar-with-dependencies.jar)

Déplacez le dans le répertoire dans un répertoire par exemple "AndroidCallLogStats".
Dans ce répertoire créez un répertoire nommé "Fichiers CLBU" (par exemple).
Créez également un répertoire nommé "Fichiers EXCEL" (par exemple).

Démarrez une ligne de commande et allez dans le répertoire "AndroidCallLogStats".

Lancez la commande suivante : 

    java -jar -DclbuFolder="./Fichiers CLBU" -DresultFolder="./Fichiers EXCEL" AndroidCallLogStats-1.0-SNAPSHOT-jar-with-dependencies.jar

Le paramètre **clbuFolder=** correspond au répertoire ou se trouvent les fichiers ".clbu" à analyser.
Le paramètre **resultFolder=** correspond au répertoire ou va être généré le fichier EXCEL.

Exemple de nom de fichier EXCEL généré : "Statistiques des appels au 2020-12-01 généré le 2020-12-03_23-34-26.xlsx".
La date "2020-12-01" correspond à la date de l'appel le plus récent trouvé dans les fichiers.

Il faut que l'environnment JAVA soit installé sur votre ordinateur pour que cette petite application fonctionne.
Vous pouvez télécharger le logiciel Java ici : [Télécharger le logiciel Java](https://www.java.com/fr/download/).

Depuis la ligne de commande vous pouvez vérifier que Java fonctionne en tapant la commande :

    java -version

Si cela fonctionne vous verrez la version de Java utilisée. La version de Java utilisée doit au moins être la version 1.8.