def worsening_norm(rastr):
    
   """This function makes the mode heavier to obtain the maximum overflow

    Parameters:
    rastr (rastr): COM object

   """

    kd = rastr.step_ut("i")
    while kd == 0:
        kd = rastr.step_ut("z")
              

def worsening_U(rastr, percent):
    
   """This function makes the mode heavier until the voltages go beyond the set limits.

    Parameters:
    rastr (rastr): COM object
    percent (int): Sets the stock percentage

   """   
    
    node_table = rastr.Tables('node')
    kd2 = rastr.step_ut("i")
    while (kd2 == 0):
        for u in range(0, node_table.size):
            if node_table.Cols('vras').Z(u) < node_table.Cols('uhom').Z(u)*0.7*percent:
                return
        kd2 = rastr.step_ut("z")

         
            
def worsening_I(rastr, i_dop):
    
   """This function makes the mode heavier until the currents go beyond the set limits.

    Parameters:
    rastr (rastr): COM object
    i_dop (string): Column name with specified current limit

   """
    
    vetv_table = rastr.Tables('vetv')
    kd3 = rastr.step_ut("i")
    while (kd3 == 0):
        for i in range(0,vetv_table.size):
            if (vetv_table.Cols(i_dop).Z(i) !=0) and (vetv_table.Cols('ib').Z(i) > \
                                                      vetv_table.Cols(i_dop).Z(i) or \
                                                      vetv_table.Cols('ie').Z(i) > vetv_table.Cols(i_dop).Z(i)):
                return
        kd3 = rastr.step_ut("z")
           

def faults(rastr, faults, shbl3):
    
   """This function disconnects the branch according to the faults file

    Parameters:
    rastr (rastr): COM object
    faults[k] (string): name of fault
    shbl3 (string): template for creating a Rastrwin3 file
    
    Return:
    j (int): line to be disconnected

   """
    
    rastr.Load(3, r'C:\Users\otrok\Downloads\regime (2).rg2', shbl3)
    vetv_table = rastr.Tables('vetv')
    for j in range(0, vetv_table.size):
        if (faults['ip'] == vetv_table.Cols('ip').Z(j)) and \
       (faults['iq'] == vetv_table.Cols('iq').Z(j)):
            vetv_table.Cols('sta').SetZ(j, 1)
            return j

def setVector(rastr, vector):
    
    """This function sets the weighting vector

    Parameters:
    rastr (rastr): COM object
    vector (dataframe): weighting vector

   """
    
    ut_table = rastr.Tables('ut_node')
    ut_table.size = len(vector)
    for index in range(0,len(vector)):
        ut_table.Cols('ny').SetZ(index, vector['node'][index])
        ut_table.Cols('tg').SetZ(index, vector['tg'][index])
        if vector['variable'][index] == 'pn':
            ut_table.Cols('pn').SetZ(index, vector['value'][index])
        else:
            ut_table.Cols('pg').SetZ(index, vector['value'][index])

def setFlowgate(rastr, flowgate):
    
    """This function sets the flowgate

    Parameters:
    rastr (rastr): COM object
    flowgate (dict): weighting vector

   """
    
    grline_table = rastr.Tables('grline')
    sechen_table = rastr.Tables('sechen')
    grline_table.size = len(flowgate)
    sechen_table.size = 1
    sechen_table.Cols('ns').SetZ(0, 1)
    for index, k in enumerate(flowgate.keys()):
        grline_table.Cols('ip').SetZ(index, flowgate[k]['ip'])
        grline_table.Cols('iq').SetZ(index, flowgate[k]['iq'])
    grline_table.Cols('ns').SetZ(0, 1)
    rastr.rgm('')




