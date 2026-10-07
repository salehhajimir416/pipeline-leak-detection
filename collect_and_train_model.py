import win32com.client
import numpy as np
import pickle
from sklearn.preprocessing import StandardScaler
from xgboost import XGBClassifier
import os
import logging
import time

# Configure logging
log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data_collection_and_model_training.log')
logging.basicConfig(filename=log_file_path, level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def access_hysys_simulation(simulation_file_name):
    try:
        logging.info(f"Connecting to HYSYS and opening simulation: {simulation_file_name}")
        hysys = win32com.client.Dispatch("HYSYS.Application")
        folder_path = os.path.dirname(os.path.abspath(__file__))
        simulation_file_path = os.path.join(folder_path, simulation_file_name)
        simcase = hysys.SimulationCases.Open(simulation_file_path)
        simcase.Activate()
        logging.info(f"Successfully opened and activated the simulation: {simulation_file_name}")
        return simcase
    except Exception as e:
        logging.error(f"Error accessing HYSYS simulation: {e}")
        return None

def get_mass_flows_from_hysys(case, stream_names):
    try:
        logging.debug("Retrieving mass flows from HYSYS...")
        flows = []
        for name in stream_names:
            stream = case.Flowsheet.MaterialStreams.Item(name)
            mass_flow = stream.MassFlow.GetValue("kg/h")
            flows.append(mass_flow)
            logging.debug(f"Stream {name} mass flow: {mass_flow} kg/h")
        logging.info(f"Retrieved mass flows: {flows}")
        return flows
    except Exception as e:
        logging.error(f"Error retrieving mass flows: {e}")
        return []

def determine_leakage_label(flows, section_pairs):
    """
    Determines if a leakage has occurred by comparing mass flows between sections.
    A leakage is assumed if the mass flow of a section is significantly lower than the previous one.

    :param flows: List of mass flows
    :param section_pairs: List of tuples, each containing the indices of streams to compare between sections
    :return: 1 if leakage detected, 0 otherwise
    """
    for current_section, previous_section in section_pairs:
        current_flow = np.mean([flows[i - 1] for i in current_section])  # Adjusting for 1-based indexing
        previous_flow = np.mean([flows[i - 1] for i in previous_section])

        # Detect leakage if the current flow is significantly lower than the previous section
        if current_flow < previous_flow * 0.95:  # Adjust the threshold as needed
            return 1  # Leakage detected
    return 0  # No leakage detected

def collect_data(case, stream_names, section_pairs, num_samples=20):
    try:
        logging.info("Collecting data from HYSYS...")
        data = []
        labels = []
        for _ in range(num_samples):
            flows = get_mass_flows_from_hysys(case, stream_names)
            if flows:
                data.append(flows)
                label = determine_leakage_label(flows, section_pairs)
                labels.append(label)
            time.sleep(5)  # Adjust sleep time as needed
        return np.array(data), np.array(labels)
    except Exception as e:
        logging.error(f"Error collecting data: {e}")
        return np.array([]), np.array([])

def save_data_to_file(data, labels, file_path):
    np.save(file_path, {'data': data, 'labels': labels})
    logging.info(f"Data and labels saved to {file_path}")

def load_handmade_model():
    try:
        model_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'xgboost_model.pkl')
        logging.info(f"Loading the handmade model from {model_file_path}...")
        with open(model_file_path, 'rb') as f:
            model_data = pickle.load(f)
        logging.info("Successfully loaded the handmade model.")
        return model_data['model'], model_data['scaler']
    except Exception as e:
        logging.error(f"Error loading the handmade model: {e}")
        return None, None

def train_and_save_model(data, labels, model=None, scaler=None):
    try:
        logging.info("Training model...")

        # Ensure labels are either 0 or 1
        unique_labels = np.unique(labels)
        if len(unique_labels) != 2 or not np.array_equal(unique_labels, [0, 1]):
            logging.error(f"Invalid classes in labels: {unique_labels}. Expected [0, 1].")
            return

        if scaler is None:
            scaler = StandardScaler()
        data_scaled = scaler.fit_transform(data)

        if model is None:
            model = XGBClassifier(use_label_encoder=False, eval_metric='logloss')
        model.fit(data_scaled, labels)

        model_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'xgboost_model.pkl')
        with open(model_file_path, 'wb') as f:
            pickle.dump({'model': model, 'scaler': scaler}, f)

        logging.info(f"Model and scaler saved to {model_file_path}")
    except Exception as e:
        logging.error(f"Error training or saving model: {e}")

if __name__ == "__main__":
    simulation_file_name = "Leakage5.1.hsc"
    stream_names = ['1', '2', '3', '4', '5', '6', '7', '8']
    
    # Define sections: compare each section with the previous one
    section_pairs = [
        ([3, 4], [1, 2]),  # Compare section 2 with section 1
        ([5, 6], [3, 4]),  # Compare section 3 with section 2
        ([7, 8], [5, 6]),  # Compare section 4 with section 3
    ]
    
    data_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'collected_data.npy')
    
    case = access_hysys_simulation(simulation_file_name)
    
    if case:
        model, scaler = load_handmade_model()
        while True:
            data, labels = collect_data(case, stream_names, section_pairs, num_samples=20)
            if data.size > 0 and labels.size > 0:
                save_data_to_file(data, labels, data_file_path)
                train_and_save_model(data, labels, model, scaler)
            else:
                logging.error("Data collection failed. Cannot save data.")
            time.sleep(60)  # Adjust sleep time as needed
    else:
        logging.error("Failed to access HYSYS simulation. Cannot proceed with data collection.")
