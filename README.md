# train_model


#  Adversarial Robustness Challenge 


##  Objectifs

- Entraîner un classifieur CNN pour reconnaître des chiffres manuscrits
- Générer des exemples adversaires avec **FGSM**
- Réentraîner le modèle avec **adversarial training**
- Évaluer la robustesse du modèle
- Visualiser les effets des attaques


##  Données utilisées

Dataset : [🔗 Kaggle - Digit Recognizer](https://www.kaggle.com/competitions/digit-recognizer/data)

Basé sur **MNIST**



## Visualisation (exemples)

Le notebook compare :
- L’image propre (avec prédiction correcte)
- L’image perturbée (avec prédiction erronée)
- En couleur BLEUE si la prédiction est correcte, ROUGE sinon



##  Reproduire l’expérience

###  Environnement

- Kaggle Notebook (GPU activé)
- TensorFlow ≥ 2.14
- Python ≥ 3.10

###  Étapes

```bash
# Cloner le projet
git clone https://github.com/gatienseven7/train_model.git
cd adversarial-mnist

# Lancer le notebook sur Kaggle (ou Google Colab)


## Ce que vous apprendrez

- Comment les modèles de ML peuvent être vulnérables
- Comment générer des attaques simples (FGSM)
- Comment rendre un modèle plus robuste face à ces attaques
- Visualisation et interprétation des failles

