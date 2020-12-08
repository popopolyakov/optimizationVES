import sys
import pfpy
import os
sys.path.append('.')
import matplotlib.pyplot as plt

folder_path = os.path.abspath('./')

models = pfpy.models

outputs = {
    'ElmRes' : 'DemoSimulation',
    'variables': {
        'SourcePCC.ElmVac': ['m:Psum:bus1','m:Qsum:bus1','m:u:bus1','m:phiu1:bus1']
    }
}

generate = True
if generate:
    # Defining model inputs
    inputpath = os.path.join(folder_path, 'signal_PCCVoltFile.csv')
    inputs = [
            # VOLTAGE INPUT
            {'network' : 'MV_EUR',
            'ElmFile' : 'PCCVolt',
            'source' : 'generate',
            'filepath' : inputpath,
            'wave_specs' : {
                    'type' : 'dip', 'tstart' : 0,
                    'tstop' : 10,   'step' : 0.01,
                    'deltat' : 0.1, 'deltay' : 0.2,
                    'y0' : 1,       't0' : 1}
            },
            # FREQUENCY INPUT
            {'network' : 'MV_EUR',
            'ElmFile' : 'PCCFreq',
            'source' : 'generate',
            # 'filepath' : filepath can be also left undefined
            'wave_specs' : {
                'type' : 'const',
                'y0' : 1,
                'tstart' : 0,
                'tstop' : 10,
                'step' : 0.01}
            }
        ]
else:
    input_path_frequency = os.path.abspath('./examples/data/input_signals/pf_freq.csv')
    input_path_voltage = os.path.abspath('./examples/data/input_signals/pf_volt.csv')

    inputs = [
            {'network'   : 'MV_EUR', # VOLTAGE INPUT
            'ElmFile'    : 'PCCVolt',
            'source'     : 'ext_file',
            'filepath'   : input_path_voltage,
        },
        {'network'    : 'NET_WITH_ELMFILE', # FREQUENCY INPUT
        'ElmFile'    : 'PCCFreq',
        'source'     : 'ext_file',
        'filepath'   : input_path_frequency
        }
    ]

# Instantiating the model and simulating
model = models.PowerFactoryModel(project_name = 'Kenigsberg',
                                       study_case='01 - Base Case', networks=['MV_EUR'],
                                       outputs = outputs, inputs = inputs)

results = model.simulate(0, 10, 0.01)