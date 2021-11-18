def worsening_norm(rastr):
    kd = rastr.step_ut("i")
    while kd == 0:
        kd = rastr.step_ut("z")
        if (kd == 0 and rastr.ut_Param(1) == 0 ) or (rastr.ut_Param(2) == 1):
            rastr.AddControl(-1,"")


def worsening_U(rastr, node_table):
    kd2 = rastr.step_ut("i")
    uk = 0
    while (kd2 == 0) and (uk == 0):
        for u in range(0,node_table.size):
            if node_table.Cols('vras').Z(u)<node_table.Cols('uhom').Z(u)*0.7*1.15:
                uk = 1
            if uk == 1:
                break
        kd2=rastr.step_ut("z")
        if (kd2 == 0 and rastr.ut_Param(1) == 0 ) or (rastr.ut_Param(2) == 1):
            rastr.AddControl(-1,"")

def worsening_I(rastr, i_dop, vetv_table):
    kd3 = rastr.step_ut("i")
    ik=0
    while (kd3 == 0) and (ik == 0):
        for i in range(0,vetv_table.size):
            if (vetv_table.Cols('ib').Z(i)>vetv_table.Cols(i_dop).Z(i) and
            vetv_table.Cols(i_dop).Z(i) !=0)  or \
            (vetv_table.Cols('ie').Z(i)>vetv_table.Cols(i_dop).Z(i) and
            vetv_table.Cols(i_dop).Z(i) !=0):
                ik = 1
            if ik == 1:
                break
        kd3=rastr.step_ut("z")
        if (kd3 == 0 and rastr.ut_Param(1) == 0 ) or (rastr.ut_Param(2) == 1):
            rastr.AddControl(-1,"")

def faults(rastr, faults, vetv_table, shbl3, k, fault):
    rastr.Load(3, r'C:\Users\otrok\Downloads\regime (2).rg2', shbl3)
    for j in range(0, vetv_table.size):
        if (faults[k]['ip'] == vetv_table.Cols('ip').Z(j)) and \
       (faults[k]['iq'] == vetv_table.Cols('iq').Z(j)):
            vetv_table.Cols('sta').SetZ(j,1)
            fault=j
            return fault
            break
    rastr.rgm('')




