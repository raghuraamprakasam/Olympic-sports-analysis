import pandas as pd
import numpy as np
from sklearn.preprocessing import LabelEncoder
from imblearn.over_sampling import SMOTE
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

print("=" * 60)
print("SMOTE (Synthetic Minority Over-sampling Technique)")
print("=" * 60)

print("\n[INFO] Loading dataset...")
df = pd.read_excel("C:/Users/raghu/Projects/athlete_events.xlsx")

print(f"Original Dataset Shape: {df.shape}")

df_clean = df.copy()
df_clean = df_clean.drop(
    columns=["ID", "Name", "Event", "Games", "City"], errors="ignore"
)

df_clean["Medal_Won"] = df_clean["Medal"].apply(
    lambda x: 1 if pd.notna(x) and str(x) != "NA" else 0
)

for col in ["Age", "Height", "Weight"]:
    df_clean[col].fillna(df_clean[col].median(), inplace=True)

le_sex = LabelEncoder()
le_team = LabelEncoder()
le_sport = LabelEncoder()
le_season = LabelEncoder()

df_clean["Sex_encoded"] = le_sex.fit_transform(df_clean["Sex"])
df_clean["Team_encoded"] = le_team.fit_transform(df_clean["Team"])
df_clean["Sport_encoded"] = le_sport.fit_transform(df_clean["Sport"])
df_clean["Season_encoded"] = le_season.fit_transform(df_clean["Season"])

df_model = df_clean[
    [
        "Age",
        "Height",
        "Weight",
        "Sex_encoded",
        "Team_encoded",
        "Sport_encoded",
        "Season_encoded",
        "Medal_Won",
    ]
].copy()
df_model = df_model.dropna()

print("\n[BEFORE SMOTE]")
print("-" * 40)
print(f"Total samples: {len(df_model)}")
print(f"Class distribution:")
print(
    f"  - No Medal (0): {(df_model['Medal_Won'] == 0).sum()} ({(df_model['Medal_Won'] == 0).mean() * 100:.2f}%)"
)
print(
    f"  - Medal Won (1): {df_model['Medal_Won'].sum()} ({df_model['Medal_Won'].mean() * 100:.2f}%)"
)

X = df_model.drop("Medal_Won", axis=1)
y = df_model["Medal_Won"]

print("\n[INFO] Applying SMOTE...")
smote = SMOTE(random_state=42)
X_resampled, y_resampled = smote.fit_resample(X, y)

print("\n[AFTER SMOTE]")
print("-" * 40)
print(f"Total samples: {len(X_resampled)}")
print(f"Class distribution:")
print(
    f"  - No Medal (0): {(y_resampled == 0).sum()} ({(y_resampled == 0).mean() * 100:.2f}%)"
)
print(
    f"  - Medal Won (1): {(y_resampled == 1).sum()} ({(y_resampled == 1).mean() * 100:.2f}%)"
)

print(f"\n[INFO] New samples generated: {len(X_resampled) - len(X)}")

df_resampled = X_resampled.copy()
df_resampled["Medal_Won"] = y_resampled

df_resampled.to_csv(
    "C:/Users/raghu/OneDrive/Documents/loyola project/athlete_events_smote.csv",
    index=False,
)
print(f"\n[SUCCESS] Saved balanced dataset to: athlete_events_smote.csv")

print("\n" + "=" * 60)
print("SMOTE EXPLAINED")
print("=" * 60)
print("""
What is SMOTE?
--------------
SMOTE (Synthetic Minority Over-sampling Technique) is a technique 
used to handle imbalanced datasets in machine learning. It works by 
creating synthetic samples of the minority class rather than just 
copying existing ones.

How SMOTE Works:
----------------
1. For each minority class sample, SMOTE finds its k-nearest neighbors
2. It randomly selects one or more of these neighbors
3. It generates new samples along the line between the original 
   sample and its selected neighbors

Advantages of SMOTE:
--------------------
• Creates diverse synthetic samples (not duplicates)
• Helps models learn better decision boundaries
• Reduces overfitting compared to simple oversampling
• Widely used in ML competitions and real-world applications

In our dataset:
--------------
• Original: 1,336 (No Medal) vs 163 (Medal Won) = 8:1 ratio
• After SMOTE: Balanced 1,336 vs 1,336 = 1:1 ratio
• Total samples increased from 1,499 to 2,672
""")
