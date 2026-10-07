# Real-Time Monitoring and Leak Detection in Fluid Pipelines

Process-simulation + ML pipeline that reads mass-flow streams from **Aspen HYSYS**, flags mass-balance drops, and classifies leak location with **XGBoost**.

## Approach

1. Open a HYSYS case via COM (`win32com`)
2. Sample mass flows on named material streams
3. Rule-based check: downstream mean flow vs upstream (threshold ~95%)
4. Scale features and classify with a trained XGBoost model

## Layout

| File | Role |
|---|---|
| `collect_and_train_model.py` | Collect simulation samples and train the classifier |
| `handmade-model.py` | Alternate / handmade-data training path |
| `real_time_monitoring.py` | Live monitoring loop, logging, and plots |
| `xgboost_model.pkl` | Trained model + scaler |
| `handmade_data.npy` | Synthetic training samples |
| `Leakage5.1.hsc` | HYSYS simulation case |
| `thesis.pdf` / `slides.pdf` | Thesis and presentation |

## Requirements

- Windows with Aspen HYSYS installed (COM automation)
- Python packages:

```bash
pip install numpy scikit-learn xgboost matplotlib tabulate pywin32
python real_time_monitoring.py
```

Logs are written next to the scripts (`real_time_monitoring.log`).
