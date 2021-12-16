import mdp_func
file_rgm=r'C:\Users\otrok\Downloads\regime (2).rg2'
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
        for i in range(0, vetv_table.size):
            idop = vetv_table.Cols(i_dop).Z(i)
            if (idop !=0) and (vetv_table.Cols('ib').Z(i) > idop or \
                               vetv_table.Cols('ie').Z(i) > idop):
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
    
    rastr.Load(3, file_rgm, shbl3)
    vetv_table = rastr.Tables('vetv')
    for j in range(0, vetv_table.size):
        if (faults['ip'] == vetv_table.Cols('ip').Z(j)) and \
       (faults['iq'] == vetv_table.Cols('iq').Z(j) and \
       faults['np']==vetv_table.Cols('np').Z(j)):
            vetv_table.Cols('sta').SetZ(j, faults[k]['sta'])
            return j

def set_vector(rastr, vector):
    
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

def set_flowgate(rastr, flowgate):
    
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

def second_criterion(rastr, faults, shbl_reg):
    
    """Second criterion

    Parameters:
    rastr (rastr): COM object
    faults (dataframe): faults
    shbl_reg (string): shablon
    
    Return:
    doavar_overflow (list): list of overflow

   """
    
    sechen_table = rastr.Tables('sechen')
    doavar_overflow = []
    for k in faults.keys():
        fault = mdp_func.faults(rastr, faults[k],  shbl_reg)
        mdp_func.worsening_norm(rastr)
        limit_overflow2=(abs(sechen_table.Cols('psech').Z(0)))
        rastr.Load(3, file_rgm, shbl_reg)
        vetv_table.Cols('sta').SetZ(fault, faults[k]['sta'])
        rastr.rgm('')
        kd = rastr.step_ut("index")
        while (kd == 0) and abs(sechen_table.Cols('psech').Z(0)) < 0.92*limit_overflow2:
            kd = rastr.step_ut("z")
        vetv_table.Cols('sta').SetZ(fault, 1-faults[k]['sta'])
        rastr.rgm('')
        doavar_overflow.append(abs(sechen_table.Cols('psech').Z(0)))
    doavar_overflow=[round(v, 2) for v in doavar_overflow]
    return doavar_overflow

def fourth_criterion(rastr, faults, shbl_reg):
    
    """Fourth criterion

    Parameters:
    rastr (rastr): COM object
    faults (dataframe): faults
    shbl_reg (string): shablon
    
    Return:
    doavar_overflow (list): list of overflow

   """
    sechen_table = rastr.Tables('sechen')
    vetv_table = rastr.Tables('vetv')
    doavar_overflow = []
    for k in faults.keys():
        fault = mdp_func.faults(rastr, faults[k], shbl_reg)
        vetv_table.Cols('sta').SetZ(fault, faults[k]['sta'])
        mdp_func.worsening_U(rastr, 1.1)
        vetv_table.Cols('sta').SetZ(fault, 1-faults[k]['sta'])
        rastr.rgm('')
        doavar_overflow.append(abs(sechen_table.Cols('psech').Z(0)))
    doavar_overflow = [round(v, 2) for v in doavar_overflow]
    return doavar_overflow

def sixth_criterion(rastr, faults, shbl_reg):
    
    """Sixth criterion

    Parameters:
    rastr (rastr): COM object
    faults (dataframe): faults
    shbl_reg (string): shablon
    
    Return:
    doavar_overflow (list): list of overflow

   """
    sechen_table = rastr.Tables('sechen')
    vetv_table = rastr.Tables('vetv')
    doavar_overflow = []
    for k in faults.keys():
        fault = mdp_func.faults(rastr, faults[k], shbl_reg)
        vetv_table.Cols('sta').SetZ(fault, faults[k]['sta'])
        mdp_func.worsening_I(rastr, 'i_dop_r_av')
        vetv_table.Cols('sta').SetZ(fault, 1-faults[k]['sta'])
        rastr.rgm('')
        doavar_overflow.append(abs(sechen_table.Cols('psech').Z(0)))
    doavar_overflow = [round(v, 2) for v in doavar_overflow]
    return doavar_overflow


