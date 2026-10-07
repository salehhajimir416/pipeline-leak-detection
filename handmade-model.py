import numpy as np
from sklearn.preprocessing import StandardScaler
from xgboost import XGBClassifier
import pickle
import os

# Define the section pairs to compare
section_pairs = [
    ([3, 4], [1, 2]),  # Compare section 2 with section 1
    ([5, 6], [3, 4]),  # Compare section 3 with section 2
    ([7, 8], [5, 6]),  # Compare section 4 with section 3
]

# Function to generate normal flow data
def generate_normal_flow():
    base_flow = np.random.uniform(30000, 40000, 4)  # Base flow for 4 sections
    flows = np.hstack([base_flow, base_flow])  # Duplicate for two streams per section
    return flows

# Function to generate leakage flow data
def generate_leakage_flow(leakage_section):
    base_flow = np.random.uniform(30000, 40000, 4)  # Base flow for 4 sections
    leakage_flow = base_flow.copy()
    leakage_flow[leakage_section - 1] *= 0.7  # Reduce flow in the leakage section
    flows = np.hstack([leakage_flow, base_flow])  # Two streams per section
    return flows

# Generate dataset
data = []
labels = []

# Generate normal operation data
for _ in range(50):
    flows = generate_normal_flow()
    data.append(flows)
    labels.append(0)  # No leakage

# Generate leakage data
for section in range(1, 4):  # Sections 2, 3, and 4
    for _ in range(10):
        flows = generate_leakage_flow(section)
        data.append(flows)
        labels.append(1)  # Leakage

data = np.array(data)
labels = np.array(labels)

# Save the handmade data for reference
data_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'handmade_data.npy')
np.save(data_file_path, {'data': data, 'labels': labels})

print(f"Handmade data saved to {data_file_path}")

# Train and save the model
def train_and_save_model(data, labels):
    try:
        print("Training model...")

        scaler = StandardScaler()
        data_scaled = scaler.fit_transform(data)

        model = XGBClassifier(use_label_encoder=False, eval_metric='logloss')
        model.fit(data_scaled, labels)

        model_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'xgboost_model.pkl')
        with open(model_file_path, 'wb') as f:
            pickle.dump({'model': model, 'scaler': scaler}, f)

        print(f"Model and scaler saved to {model_file_path}")
    except Exception as e:
        print(f"Error training or saving model: {e}")

train_and_save_model(data, labels)

