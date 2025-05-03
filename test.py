# Chargement des bibliothèques nécessaires
import matplotlib.pyplot as plt
from matplotlib_venn import venn2

# Définition des ensembles à partir des deux listes
liste1 = {
    "mindset", "habits", "habit", "students", "die", "children", "school", "com",
    "will", "life", "york", "change", "action", "mind", "said", "people", "“i",
    "you’ll", "can", "went", "didn’t", "told", "think", "learn", "wanted", "use",
    "young", "book", "came", "want", "you’re", "upon", "ability", "high", "create",
    "make", "law", "hard", "got", "without", "company", "small", "every", "day",
    "example", "need", "human", "process", "later", "personal"
}

liste2 = {
    "income", "percent", "companies", "wealth", "mr", "life", "company", "market",
    "high", "feel", "children", "parents", "research", "students", "year", "school",
    "average", "want", "power", "things", "thoughts", "journal", "business", "likely",
    "mindset", "study", "mind", "you’re", "can", "american", "group", "will", "you’ve",
    "college", "change", "body", "university", "social", "less", "love", "financial",
    "two", "york", "worth", "number", "negative", "pain", "you’ll", "person", "action"
}

# Création du diagramme de Venn
plt.figure(figsize=(8, 6))
venn = venn2([liste1, liste2], set_labels=('Liste 1', 'Liste 2'))
plt.title("Diagramme de Venn des mots des deux listes")
plt.show()