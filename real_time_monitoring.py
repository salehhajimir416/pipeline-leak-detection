import time
import win32com.client
import numpy as np
import pickle
import logging
import os
import matplotlib.pyplot as plt
from tabulate import tabulate

# Configure logging
log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'real_time_monitoring.log')
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

def detect_potential_leakage(flows, section_pairs, threshold=0.95):
    leakages_detected = []
    for section_id, (current_section, previous_section) in enumerate(section_pairs, start=1):
        current_flow = np.mean([flows[i - 1] for i in current_section])
        previous_flow = np.mean([flows[i - 1] for i in previous_section])

        logging.debug(f"Current section flow: {current_flow}, Previous section flow: {previous_flow}")
        
        if current_flow < previous_flow * threshold:
            logging.info(f"Potential leakage detected in Section {section_id}: "
                         f"Current Flow: {current_flow}, Previous Flow: {previous_flow}")
            leakages_detected.append(section_id)
        else:
            logging.info(f"No leakage detected in Section {section_id}: "
                         f"Current Flow: {current_flow}, Previous Flow: {previous_flow}")

    return leakages_detected

def load_trained_model():
    try:
        model_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'xgboost_model.pkl')
        with open(model_file_path, 'rb') as f:
            model_data = pickle.load(f)
            model = model_data['model']
            scaler = model_data['scaler']
            logging.info("Loaded trained model successfully.")
            return model, scaler
    except Exception as e:
        logging.error(f"Error loading model: {e}")
        return None, None

def classify_with_model(model, scaler, flows):
    try:
        data_scaled = scaler.transform([flows])
        prediction = model.predict(data_scaled)[0]
        logging.info(f"Model prediction: {prediction}")
        return prediction
    except Exception as e:
        logging.error(f"Error during model classification: {e}")
        return None

def report_leakages(leakage_sections, timestamp, flows, section_pairs):
    headers = ['Section', 'Leakage Detected', 'Flow Difference']
    table_data = []
    for section_id in leakage_sections:
        current_section, previous_section = section_pairs[section_id - 1]
        current_flow = np.mean([flows[i - 1] for i in current_section])
        previous_flow = np.mean([flows[i - 1] for i in previous_section])
        flow_diff = previous_flow - current_flow
        table_data.append((f"Section {section_id}", "Yes", f"{flow_diff:.2f} kg/h"))

    report_table = tabulate(table_data, headers, tablefmt="grid")
    print(f"\n[ALERT] Leakages detected at {timestamp}:\n{report_table}")
    logging.info(f"Leakage detected at {timestamp} in sections: {leakage_sections}")

def plot_all_streams(flows, section_pairs, threshold=0.95):
    """
    Plot the mass flow of all streams with bars and threshold lines.
    """
    plt.figure(figsize=(12, 6))
    stream_names = ['1', '2', '3', '4', '5', '6', '7', '8']  # Adjust this if you have more streams
    bar_width = 0.4
    index = np.arange(len(stream_names))

    # Plot mass flows as bars
    plt.bar(index, flows, width=bar_width, color='b', label='Mass Flow')

    # Colors for different thresholds
    threshold_colors = ['r', 'g', 'orange']  # Add more colors if needed

    # Calculate and plot threshold lines for each stream
    for section_id, (current_section, previous_section) in enumerate(section_pairs, start=1):
        previous_flow = np.mean([flows[i - 1] for i in previous_section])
        threshold_flow = previous_flow * threshold
        plt.axhline(y=threshold_flow, color=threshold_colors[section_id - 1], linestyle='--', 
                    label=f'Threshold for Section {section_id}')

    plt.title('Mass Flow of All Streams in the Pipeline')
    plt.xlabel('Stream Number')
    plt.ylabel('Mass Flow (kg/h)')
    plt.xticks(index, stream_names)
    plt.legend(title='Legend', loc='upper right')
    plt.grid(axis='y')
    plt.tight_layout()
    plt.show()

def real_time_monitoring(case, stream_names, section_pairs, threshold=0.95):
    model, scaler = load_trained_model()
    if model is None or scaler is None:
        logging.error("Model is not loaded. Exiting real-time monitoring.")
        return

    try:
        while True:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            flows = get_mass_flows_from_hysys(case, stream_names)
            if flows:
                potential_leakages = detect_potential_leakage(flows, section_pairs, threshold=threshold)

                if potential_leakages:
                    prediction = classify_with_model(model, scaler, flows)
                    if prediction == 1:
                        report_leakages(potential_leakages, timestamp, flows, section_pairs)
                    else:
                        logging.info(f"No leakage detected by the model at {timestamp}.")
                else:
                    logging.info(f"No potential leakage detected at {timestamp}.")

                # Always plot all streams after checking for leakages
                plot_all_streams(flows, section_pairs, threshold)

            else:
                logging.error("No mass flow data available. Skipping this iteration.")

            time.sleep(10)  # Adjust for real-time monitoring frequency
    except Exception as e:
        logging.error(f"Error during real-time monitoring: {e}")

if __name__ == "__main__":
    simulation_file_name = "Leakage5.1.hsc"
    stream_names = ['1', '2', '3', '4', '5', '6', '7', '8']

    section_pairs = [
        ([3, 4], [1, 2]),  # Compare section 1
        ([5, 6], [3, 4]),  # Compare section 2
        ([7, 8], [5, 6]),  # Compare section 3
    ]

    case = access_hysys_simulation(simulation_file_name)
    if case:
        real_time_monitoring(case, stream_names, section_pairs, threshold=0.95)
