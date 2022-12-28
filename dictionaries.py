import datetime
#ROOT = os.path.dirname(__file__)

SQL = {
    'LESN': "SELECT F_ARODES.ADRESS_FOREST, F_ARODES.ARODES_INT_NUM, F_FOREST_RANGE.FOREST_RANGE_NAME FROM F_ARODES "
            "INNER JOIN F_FOREST_RANGE ON F_ARODES.ARODES_INT_NUM = F_FOREST_RANGE.ARODES_INT_NUM WHERE ((("
            "F_ARODES.ARODES_TYP_CD)='L-CTWO'))",
    'OBR': "SELECT F_ARODES.ADRESS_FOREST FROM F_ARODES WHERE (((F_ARODES.ARODES_TYP_CD)='OBRĘB'))",
    'WYDZ': """SELECT F_ARODES.TEMP_ADRESS_FOREST FROM F_ARODES WHERE (((F_ARODES.ARODES_TYP_CD) Like 'WYDZIEL'));"""

    }
SQL2 = 'SELECT F_COMMUNITY.COMMUNITY_CD, F_COMMUNITY.COMMUNITY_NAME FROM F_COMMUNITY'

RAPORTS = {
        'Lista Kontrolna':
            {
             'Błędy': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_ERROR_DIC.ERROR_MESSAGE, F_SUBAREA.RECONSTR_CD
FROM F_ERROR_DIC INNER JOIN ((F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN (F_ERROR_HEAD INNER JOIN F_ERROR_LOG ON F_ERROR_HEAD.ARODES_INT_NUM = F_ERROR_LOG.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_ERROR_HEAD.ARODES_INT_NUM) ON F_ERROR_DIC.ERROR_CD = F_ERROR_LOG.ERROR_CD
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ERROR_DIC.ERROR_CD) Not In ('pro01')) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
ORDER BY F_ARODES.ORDER_KEY, F_ARODES.ORDER_KEY;
""",False, 'BŁĘDY', ],
                'Uszkodzenia': ["""SELECT F_SUBAREA.CAUSE_CD, F_SUBAREA.DAMAGE_DEGREE, F_SUBAREA.SUB_AREA, F_ARODES.TEMP_ADRESS_FOREST, F_AROD_PEST.PEST_CD, F_PEST_DIC.PEST_NR, F_PEST_DIC.PEST_NAME, GAT_PAN.SPECIES_CD, GAT_PAN.SPECIES_AGE, F_SUBAREA.SUBAREA_INFO
FROM F_PEST_DIC RIGHT JOIN (F_ARODES INNER JOIN ((F_AROD_PEST RIGHT JOIN F_SUBAREA ON F_AROD_PEST.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN GAT_PAN ON F_SUBAREA.ARODES_INT_NUM = GAT_PAN.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_PEST_DIC.PEST_CD = F_AROD_PEST.PEST_CD
GROUP BY F_SUBAREA.CAUSE_CD, F_SUBAREA.DAMAGE_DEGREE, F_SUBAREA.SUB_AREA, F_ARODES.TEMP_ADRESS_FOREST, F_AROD_PEST.PEST_CD, F_PEST_DIC.PEST_NR, F_PEST_DIC.PEST_NAME, GAT_PAN.SPECIES_CD, GAT_PAN.SPECIES_AGE, F_SUBAREA.SUBAREA_INFO, F_ARODES.ORDER_KEY, F_ARODES.TEMP_ACT_ADRESS
HAVING (((F_SUBAREA.DAMAGE_DEGREE)>0) AND ((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
ORDER BY F_ARODES.TEMP_ADRESS_FOREST, F_ARODES.ORDER_KEY;""", False, 'USZKODZENIA', ],
                'Jakie Cechy Drzewostanu': ["""SELECT F_AROD_STAND_PEC.FOREST_PEC_CD
FROM F_ARODES INNER JOIN F_AROD_STAND_PEC ON F_ARODES.ARODES_INT_NUM = F_AROD_STAND_PEC.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
GROUP BY F_AROD_STAND_PEC.FOREST_PEC_CD;
""", False, 'JAKIE_CECHA', ],
                'Pusta Cecha Drzewostanu': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_AROD_STAND_PEC.FOREST_PEC_CD, F_SUBAREA.AREA_TYPE_CD, F_SUBAREA.ARODES_INT_NUM
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) LEFT JOIN F_AROD_STAND_PEC ON F_SUBAREA.ARODES_INT_NUM = F_AROD_STAND_PEC.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_STAND_PEC.FOREST_PEC_CD) Is Null) AND ((F_SUBAREA.AREA_TYPE_CD)='d-stan') AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
""", False, 'PUSTA CECHA',],
                'Rodzaje Powierzchni':["""SELECT F_SUBAREA.AREA_TYPE_CD, F_AREA_TYPE_DIC.AREA_TYPE_NR
FROM F_ARODES INNER JOIN (F_AREA_TYPE_DIC INNER JOIN F_SUBAREA ON F_AREA_TYPE_DIC.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
GROUP BY F_SUBAREA.AREA_TYPE_CD, F_AREA_TYPE_DIC.AREA_TYPE_NR;
""", False, 'RODZ_POW', ],
                'Pusta Zgodność Siedliskowa': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.TEMP_ACC_ESTIM, F_SUBAREA.ACCORD_CD, F_SUBAREA.AREA_TYPE_CD, F_GROUP_CATEGORY.SUPERGR_CAT_NAME, F_SUBAREA.SITE_TYPE_CD
FROM F_ARODES INNER JOIN (F_GROUP_CATEGORY INNER JOIN F_SUBAREA ON F_GROUP_CATEGORY.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.TEMP_ACC_ESTIM) Is Null) AND ((F_GROUP_CATEGORY.SUPERGR_CAT_NAME)='grunty leśne zalesione'))
ORDER BY F_ARODES.TEMP_ADRESS_FOREST;
""",False, 'PUSTA_ZGODN_SIEDL', ],
                'Pusta Bonitacja': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.AREA_TYPE_CD, F_AROD_STOREY.STOREY_CD, F_STOREY_SPECIES.SPECIES_CD, F_STOREY_SPECIES.SPECIES_AGE, F_STOREY_SPECIES.PART_CD, F_STOREY_SPECIES.TEMP_STCL_ESTIM, F_STOREY_SPECIES.HEIGHT, F_SUBAREA.SITE_TYPE_CD, F_ARODES.TEMP_AVB_SAMPLE
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN (F_AROD_STOREY INNER JOIN F_STOREY_SPECIES ON (F_AROD_STOREY.ARODES_INT_NUM = F_STOREY_SPECIES.ARODES_INT_NUM) AND (F_AROD_STOREY.STOREY_CD = F_STOREY_SPECIES.STOREY_CD)) ON F_SUBAREA.ARODES_INT_NUM = F_AROD_STOREY.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.AREA_TYPE_CD)='d-stan') AND ((F_AROD_STOREY.STOREY_CD) In ('DRZEW','IP','IIP')) AND ((F_STOREY_SPECIES.PART_CD) Not In ('pjd','mjs')) AND ((F_STOREY_SPECIES.TEMP_STCL_ESTIM) Is Null) AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
""",False, 'PUSTA_BONIT', ],
                'Makrorzeźba Położenie': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.AREA_TYPE_CD, F_SUBAREA.RELIEF_CD, F_SUBAREA.POSITION_CD
FROM F_ARODES INNER JOIN ((F_AREA_TYPE_DIC INNER JOIN F_GROUP_CATEGORY ON F_AREA_TYPE_DIC.AREA_TYPE_CD = F_GROUP_CATEGORY.AREA_TYPE_CD) INNER JOIN F_SUBAREA ON F_AREA_TYPE_DIC.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_GROUP_CATEGORY.SUPERGR_CAT_NAME) In ('grunty leśne zalesione','grunty leśne niezalesione')) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
ORDER BY F_SUBAREA.RELIEF_CD, F_SUBAREA.POSITION_CD;
""", False, 'MAKRORZEŹBA_POLO', ],
                'Pochodzenie Podrostu': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_AROD_SPEC_PEC.FOREST_PEC_CD
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN (F_AROD_STOREY INNER JOIN (F_AROD_SPEC_PEC RIGHT JOIN F_STOREY_SPECIES ON (F_AROD_SPEC_PEC.SPEC_STOR_INT_NUM = F_STOREY_SPECIES.SPEC_STOR_INT_NUM) AND (F_AROD_SPEC_PEC.ARODES_INT_NUM = F_STOREY_SPECIES.ARODES_INT_NUM)) ON (F_AROD_STOREY.ARODES_INT_NUM = F_STOREY_SPECIES.ARODES_INT_NUM) AND (F_AROD_STOREY.STOREY_CD = F_STOREY_SPECIES.STOREY_CD)) ON F_SUBAREA.ARODES_INT_NUM = F_AROD_STOREY.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_SPEC_PEC.FOREST_PEC_CD) Is Null) AND ((F_STOREY_SPECIES.STOREY_CD) In ('podr','podrii')) AND ((F_STOREY_SPECIES.PART_CD) Not In ('MJS','PJD')) AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
""", False, 'POCHODZENIE_PODROSTU', ],
                'Gatunki w Warstwach': ["""SELECT F_STOREY_SPECIES.SPECIES_CD, F_STOREY_SPECIES.STOREY_CD, F_TREE_SPECIES.SPECIES_NAME, F_STOREY_SPECIES.PART_CD
FROM F_TREE_SPECIES INNER JOIN (F_ARODES INNER JOIN F_STOREY_SPECIES ON F_ARODES.ARODES_INT_NUM = F_STOREY_SPECIES.ARODES_INT_NUM) ON F_TREE_SPECIES.SPECIES_CD = F_STOREY_SPECIES.SPECIES_CD
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?))
GROUP BY F_STOREY_SPECIES.SPECIES_CD, F_STOREY_SPECIES.STOREY_CD, F_TREE_SPECIES.SPECIES_NAME, F_STOREY_SPECIES.PART_CD;
""", False, 'GAT_W_WARSTW', ],
                'Subarea Info': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.SUBAREA_INFO, [skrocony_opis-taks].ot, F_SUBAREA.AREA_TYPE_CD, F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.SUB_AREA
FROM F_ARODES INNER JOIN (F_SUBAREA LEFT JOIN [skrocony_opis-taks] ON F_SUBAREA.ARODES_INT_NUM = [skrocony_opis-taks].ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.SUBAREA_INFO) Is Not Null) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
ORDER BY F_ARODES.ORDER_KEY;
""",False, 'SUBAREA_INFO', ],
                'Ekosystemy Referencyjne': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.AREA_TYPE_CD, F_SUBAREA.SUB_AREA, F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.SUBAREA_INFO, F_SUBAREA.FOREST_FUNC_CD, [skrocony_opis-taks].ot, F_AROD_CUE.MEASURE_CD, F_SUBAREA.SILVICULTURE_CD
FROM (F_ARODES INNER JOIN ([skrocony_opis-taks] RIGHT JOIN F_SUBAREA ON [skrocony_opis-taks].ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_CUE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CUE.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.SUBAREA_INFO) Like '%EKO_REF%') AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
ORDER BY F_ARODES.ORDER_KEY;
""",False,'EKO_REF',],
                'Siedliska Przyrodnicze': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.AREA_TYPE_CD, F_AROD_PROT_SITE.PROT_SITE_CD, F_AROD_PROT_SITE.PROT_SITE_STATE, F_SUBAREA.SUB_AREA, F_AROD_PROT_SITE.PROT_SITE_AREA, F_AROD_PROT_SITE.LOCATION_CD, F_SUBAREA.SUBAREA_INFO, F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.FOREST_FUNC_CD, [f_arodes].[temp_adress_forest] & '_' & [f_arod_prot_site].[prot_site_cd] AS klucz
FROM F_ARODES INNER JOIN (F_SUBAREA INNER JOIN F_AROD_PROT_SITE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_PROT_SITE.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
ORDER BY F_ARODES.ORDER_KEY;
""", False, 'SIEDL_PRZYR', ],
                'Brak Wskazówek': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.AREA_TYPE_CD, F_AROD_CUE.ARODES_INT_NUM, F_SUBAREA.SUBAREA_INFO
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) LEFT JOIN F_AROD_CUE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CUE.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.AREA_TYPE_CD)='D-STAN') AND ((F_AROD_CUE.ARODES_INT_NUM) Is Null) AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
""", False, 'BRAK_WSK',],
                'Wskazówki Nieleśne': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.AREA_TYPE_CD, F_AROD_CUE.MEASURE_CD, F_AROD_CUE.LARGE_TIMBER_PERC, F_GROUP_CATEGORY.TYPE_FL
FROM (F_ARODES INNER JOIN (F_GROUP_CATEGORY INNER JOIN F_SUBAREA ON F_GROUP_CATEGORY.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_CUE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CUE.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_GROUP_CATEGORY.TYPE_FL)='n') AND ((F_ARODES.TEMP_ACT_ADRESS)=True) AND ((F_AROD_CUE.CUE_RANK_ORDER)=1))
ORDER BY F_ARODES.ORDER_KEY;
""",False, 'WSK_NIELEŚNE', ],
                'Jakie Wskazania': ["""SELECT F_AROD_CUE.MEASURE_CD, Sum(F_AROD_CUE.CUTTING_AREA) AS SumaOfCUTTING_AREA, F_AROD_CUE.CUTTING_NR
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_CUE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CUE.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
GROUP BY F_AROD_CUE.MEASURE_CD, F_AROD_CUE.CUTTING_NR;
""", False, 'JAKIE_WSK', ],
                'Wskazówki Popr': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.SUB_AREA, F_SUBAREA.SITE_TYPE_CD, F_STOREY_SPECIES.SPECIES_CD, F_STOREY_SPECIES.SPECIES_AGE, F_AROD_STOREY.TEMP_STDEN_ESTIM, F_AROD_CUE.MEASURE_CD, F_AROD_CUE.CUTTING_AREA, F_AROD_CUE.LARGE_TIMBER_PERC, F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.ROTATION_AGE, F_SUBAREA.SUBAREA_INFO, F_AROD_PROT_SITE.PROT_SITE_CD, F_SUBAREA.RECONSTR_CD
FROM ((F_ARODES INNER JOIN (F_AROD_PROT_SITE RIGHT JOIN F_SUBAREA ON F_AROD_PROT_SITE.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_CUE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CUE.ARODES_INT_NUM) INNER JOIN (F_AROD_STOREY INNER JOIN F_STOREY_SPECIES ON (F_AROD_STOREY.ARODES_INT_NUM = F_STOREY_SPECIES.ARODES_INT_NUM) AND (F_AROD_STOREY.STOREY_CD = F_STOREY_SPECIES.STOREY_CD)) ON F_SUBAREA.ARODES_INT_NUM = F_AROD_STOREY.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_STOREY.STOREY_CD) In ('IP','DRZEW')) AND ((F_SUBAREA.AREA_TYPE_CD)='D-STAN') AND ((F_ARODES.TEMP_ACT_ADRESS)=True) AND ((F_STOREY_SPECIES.SPECIES_RANK_ORDER)=1) AND ((F_AROD_CUE.CUE_RANK_ORDER)=1))
ORDER BY F_STOREY_SPECIES.SPECIES_CD, F_STOREY_SPECIES.SPECIES_AGE, F_ARODES.ORDER_KEY, F_STOREY_SPECIES.SPECIES_CD, F_STOREY_SPECIES.SPECIES_AGE;
""", False, 'WSK_POPR',],
                'Zadrzewienie': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_AROD_STOREY.*, F_SUBAREA.STAND_STRUCT_CD, F_SUBAREA.AREA_TYPE_CD, GAT_PAN.SPECIES_CD, GAT_PAN.SPECIES_AGE
FROM (F_ARODES INNER JOIN F_AROD_STOREY ON F_ARODES.ARODES_INT_NUM = F_AROD_STOREY.ARODES_INT_NUM) INNER JOIN (GAT_PAN INNER JOIN F_SUBAREA ON GAT_PAN.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON (F_SUBAREA.ARODES_INT_NUM = F_AROD_STOREY.ARODES_INT_NUM) AND (F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM)
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.AREA_TYPE_CD)='d-stan') AND ((F_AROD_STOREY.STOREY_CD) In ('drzew','ip')))
ORDER BY F_ARODES.ORDER_KEY;
""", False,'ZADRZEW', ],
                'Okres Odnowienia': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.STAND_STRUCT_CD, F_SUBAREA.SUB_AREA, [skrocony_opis-taks].ot, F_SUBAREA.TEMP_ACC_ESTIM, F_SUBAREA.RECONSTR_PERIOD, F_AROD_CUE.MEASURE_CD, F_SUBAREA.RECONSTR_CD
FROM (F_ARODES INNER JOIN ([skrocony_opis-taks] RIGHT JOIN F_SUBAREA ON [skrocony_opis-taks].ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_CUE ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CUE.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.RECONSTR_PERIOD) Is Not Null) AND ((F_ARODES.TEMP_ACT_ADRESS)=True) AND ((F_AROD_CUE.CUE_RANK_ORDER)=1));
""", False, 'OKRES_ODNO',],
                'TSL Uwilgotnienie': ["""SELECT F_SUBAREA.SITE_TYPE_CD, F_SUBAREA.MOISTURY_CD
FROM F_MOISTENING_DIC INNER JOIN (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_MOISTENING_DIC.MOISTURY_CD = F_SUBAREA.MOISTURY_CD
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
GROUP BY F_SUBAREA.SITE_TYPE_CD, F_SUBAREA.MOISTURY_CD;
""", False, 'TSL_UWILG', ],
                'Porolne Zgodność': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.SITE_TYPE_CD, F_SUBAREA.TEMP_ACC_ESTIM, F_SUBAREA.DEGRADATION_CD, F_SUBAREA.SOIL_PEC_CD
FROM F_ARODES INNER JOIN (F_GROUP_CATEGORY INNER JOIN F_SUBAREA ON F_GROUP_CATEGORY.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.TEMP_ACC_ESTIM)='niezg') AND ((F_SUBAREA.SOIL_PEC_CD)='POROL') AND ((F_ARODES.TEMP_ACT_ADRESS)=True) AND ((F_GROUP_CATEGORY.SUPERGR_CAT_NAME)='grunty leśne zalesione'));
""", False, 'POROLNE_ZGOD',],
                'Porol_1':["""SELECT POROL_1Bt.TEMP_ADRESS_FOREST, POROL_1Bt.FOREST_PEC_CD, POROL_1At.TEMP_ADRESS_FOREST, POROL_1At.SOIL_PEC_CD, POROL_1At.AREA_TYPE_CD, GAT_PAN.SPECIES_CD, GAT_PAN.SPECIES_AGE
FROM GAT_PAN INNER JOIN (POROL_1Bt RIGHT JOIN POROL_1At ON POROL_1Bt.TEMP_ADRESS_FOREST = POROL_1At.TEMP_ADRESS_FOREST) ON GAT_PAN.ARODES_INT_NUM = POROL_1At.ARODES_INT_NUM
WHERE (((POROL_1Bt.FOREST_PEC_CD) Is Null) AND ((POROL_1At.AREA_TYPE_CD)='D-STAN'))
ORDER BY POROL_1Bt.TEMP_ADRESS_FOREST;
""", False, 'POROL_1',], #WATCHOUT ->dodatkowe query w main.py//final() wazne! bez dodatkowych nie ruszy
                'Porol_2':["""SELECT F_ARODES.ARODES_INT_NUM, POROL_1Bt.TEMP_ADRESS_FOREST, POROL_1At.TEMP_ADRESS_FOREST
FROM F_ARODES INNER JOIN (POROL_1Bt LEFT JOIN POROL_1At ON POROL_1Bt.TEMP_ADRESS_FOREST = POROL_1At.TEMP_ADRESS_FOREST) ON F_ARODES.TEMP_ADRESS_FOREST = POROL_1Bt.TEMP_ADRESS_FOREST
WHERE (((POROL_1At.TEMP_ADRESS_FOREST) Is Null) AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
""", False, 'POROL_2',], #WATCHOUT ->dodatkowe query w main.py//final() wazne! bez dodatkowych nie ruszy
                'TSL_Gospodarstwa':["""SELECT F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.SITE_TYPE_CD
FROM F_ARODES INNER JOIN (F_GROUP_CATEGORY INNER JOIN F_SUBAREA ON F_GROUP_CATEGORY.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_GROUP_CATEGORY.SUPERGR_CAT_NAME) In ('grunty leśne zalesione')))
GROUP BY F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.SITE_TYPE_CD;
""",False, 'TSL_GOSP',],
                'Zlożenie GTD': ["""SELECT GOAL1t.TEMP_ADRESS_FOREST, GOAL1t.SITE_TYPE_CD, F_SUBAREA.MOISTURY_CD, F_SUBAREA.SOIL_SUBTYPE_CD, [GOAL1t].[SPECIES_CD] & [GOAL2t].[SPECIES_CD] & [GOAL3t].[SPECIES_CD] & [GOAL4t].[SPECIES_CD] AS GTD, F_SUBAREA.SILVICULTURE_CD, F_SUBAREA.SUB_AREA, F_SUBAREA.SUBAREA_INFO, F_SUBAREA.TEMP_ACC_ESTIM, F_SUBAREA.SOIL_PEC_CD, F_SUBAREA.AREA_TYPE_CD, F_SUBAREA.ARODES_INT_NUM, F_SITE_TYPE_DIC.SITE_TYPE_NR
FROM F_SITE_TYPE_DIC INNER JOIN (F_ARODES INNER JOIN ((F_SUBAREA INNER JOIN ((GOAL1t LEFT JOIN GOAL2t ON GOAL1t.ARODES_INT_NUM = GOAL2t.ARODES_INT_NUM) LEFT JOIN GOAL3t ON GOAL2t.ARODES_INT_NUM = GOAL3t.ARODES_INT_NUM) ON F_SUBAREA.ARODES_INT_NUM = GOAL1t.ARODES_INT_NUM) LEFT JOIN GOAL4t ON GOAL3t.ARODES_INT_NUM = GOAL4t.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_SITE_TYPE_DIC.SITE_TYPE_CD = F_SUBAREA.SITE_TYPE_CD
WHERE (((GOAL1t.TEMP_ADRESS_FOREST) Like ?))
ORDER BY F_ARODES.ORDER_KEY;
""", False, 'ZLOZ_GTD1',]
             },
    }



def create_title_page_dict(typ, community, area, act_date, start_date, end_date, county, district, municipality, nadle):
    title_dic = {
        '//type_report//': typ,
        '//COMMUNITY_NAME//': community,
        '//pow//': area,
        '//stan_na_dzien//': act_date,
        '//okres_od//': start_date,
        '//okres_do//': end_date,
        '//county_name//': county,
        '//district_name//': district,
        '//municipality_name//': municipality,
        '//nadlesnictwo//': nadle,
        '//AKTUAL_DATE//': str(datetime.date.today().strftime('%d/%m/%Y'))
    }
    return title_dic


def create_adr_for_dict(full_adr, list=False):
    if list:
        address_dic = {
            'RDLP': [],
            'NADL': [],
            'OBR': [],
            'LCTWO': [],
            'ODDZ': [],
            'PODODDZ': [],
            'WYDZ': []
        }

        for adr in full_adr:
            address_dic['RDLP'].append(adr[0:2]),
            address_dic['NADL'].append(adr[3:5]),
            address_dic['OBR'].append(adr[6]),
            address_dic['LCTWO'].append(adr[8:10]),
            address_dic['ODDZ'].append(adr[11:17]),
            address_dic['PODODDZ'].append(adr[18:22]),
            address_dic['WYDZ'].append(adr[23:25])
        # for key in address_dic:
        #     if key in ['RDLP', 'NADL', 'OBR', 'LCTWO']:
        #         list = address_dic[key]
        #         removed = remove_dups(list)
        #         address_dic[key] = removed
    else:
        address_dic = {
            'RDLP': full_adr[0:2],
            'NADL': full_adr[3:5],
            'OBR': full_adr[6],
            'LCTWO': full_adr[8:10],
            'ODDZ': full_adr[11:17],
            'PODODDZ': full_adr[18:22],
            'WYDZ': full_adr[23:25]
        }
    return address_dic


def remove_dups(x):
    return list(set(x))


def raport_Dict(): # metoda do dynamicznego przypisywania adresow i innych zmiennych do kwerend. Nadal bedize
    # mozna zrobic create_raport z maina. tlyko trzeba wywolac te funcke przed tworzeniem raportow

    RAPORTS = {
        'Raporty':
            {'Rozbieżności':["""SELECT [COUNTY_CD] & '-' & [DISTRICT_CD] & '-' & [MUNICIPALITY_CD] & '-' & [COMMUNITY_CD] AS ADDR, F_ARODES.TEMP_ADRESS_FOREST, F_PARCEL.PARCEL_NR, F_SUBAREA.AREA_TYPE_CD, F_PARCEL_LAND_USE.AREA_USE_CD AS UZ, F_PARCEL_LAND_USE.LAND_USE_AREA, F_GROUP_CATEGORY.AREA_USE_CD, F_AROD_LAND_USE.AROD_LAND_USE_AREA
FROM F_PARCEL_LAND_USE INNER JOIN ((F_ARODES INNER JOIN (F_PARCEL INNER JOIN F_AROD_LAND_USE ON F_PARCEL.PARCEL_INT_NUM = F_AROD_LAND_USE.PARCEL_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_AROD_LAND_USE.ARODES_INT_NUM) INNER JOIN (F_GROUP_CATEGORY INNER JOIN F_SUBAREA ON F_GROUP_CATEGORY.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON (F_PARCEL.PARCEL_INT_NUM = F_PARCEL_LAND_USE.PARCEL_INT_NUM) AND (F_PARCEL_LAND_USE.SHAPE_NR = F_AROD_LAND_USE.SHAPE_NR) AND (F_PARCEL_LAND_USE.PARCEL_INT_NUM = F_AROD_LAND_USE.PARCEL_INT_NUM)
WHERE (((F_ARODES.TEMP_ACT_ADRESS)=True))
GROUP BY [COUNTY_CD] & '-' & [DISTRICT_CD] & '-' & [MUNICIPALITY_CD] & '-' & [COMMUNITY_CD], F_ARODES.TEMP_ADRESS_FOREST, F_PARCEL.PARCEL_NR, F_SUBAREA.AREA_TYPE_CD, F_PARCEL_LAND_USE.AREA_USE_CD, F_PARCEL_LAND_USE.LAND_USE_AREA, F_GROUP_CATEGORY.AREA_USE_CD, F_AROD_LAND_USE.AROD_LAND_USE_AREA, F_PARCEL.COUNTY_CD, F_PARCEL.DISTRICT_CD, F_PARCEL.MUNICIPALITY_CD, F_PARCEL.COMMUNITY_CD, F_PARCEL.PARCEL_INT_NUM, F_ARODES.ORDER_KEY, [F_ARODES].[ADRESS_FOREST] & '-' & [F_PARCEL_LAND_USE].[PARCEL_INT_NUM] & '-' & [F_PARCEL_LAND_USE].[SHAPE_NR], F_GROUP_CATEGORY.TYPE_FL, F_PARCEL_LAND_USE.AREA_USE_CD, F_PARCEL_LAND_USE.SHAPE_NR
HAVING (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_GROUP_CATEGORY.TYPE_FL)='l') AND ((F_PARCEL_LAND_USE.AREA_USE_CD)<>'ls'))

UNION ALL SELECT [COUNTY_CD] & '-' & [DISTRICT_CD] & '-' & [MUNICIPALITY_CD] & '-' & [COMMUNITY_CD] AS ADDR, F_ARODES.TEMP_ADRESS_FOREST, F_PARCEL.PARCEL_NR, F_SUBAREA.AREA_TYPE_CD, F_PARCEL_LAND_USE.AREA_USE_CD AS UZ, F_PARCEL_LAND_USE.LAND_USE_AREA, F_GROUP_CATEGORY.AREA_USE_CD, F_AROD_LAND_USE.AROD_LAND_USE_AREA
FROM F_PARCEL INNER JOIN (F_PARCEL_LAND_USE INNER JOIN ((F_ARODES INNER JOIN F_AROD_LAND_USE ON F_ARODES.ARODES_INT_NUM = F_AROD_LAND_USE.ARODES_INT_NUM) INNER JOIN (F_SUBAREA INNER JOIN F_GROUP_CATEGORY ON F_SUBAREA.AREA_TYPE_CD = F_GROUP_CATEGORY.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON (F_PARCEL_LAND_USE.SHAPE_NR = F_AROD_LAND_USE.SHAPE_NR) AND (F_PARCEL_LAND_USE.PARCEL_INT_NUM = F_AROD_LAND_USE.PARCEL_INT_NUM)) ON F_PARCEL.PARCEL_INT_NUM = F_PARCEL_LAND_USE.PARCEL_INT_NUM
WHERE (((F_GROUP_CATEGORY.AREA_USE_CD)<>[f_parcel_land_use].[area_use_cd]) AND ((F_GROUP_CATEGORY.TYPE_FL)='n') AND ((F_ARODES.TEMP_ACT_ADRESS)=True)) AND F_ARODES.TEMP_ADRESS_FOREST LIKE ?;
""", False, 'ROZB', ], #achtung
             'Naruszenia stanu posiadania': ["""SELECT * FROM PISMO_NARUSZENIA WHERE [Adres leśny] LIKE ?""", False, 'NA_ST_PS', ],
             'Linie energetyczne na LS': ["""SELECT * FROM PISMO_LINIE_ENERG WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'L_ENERG_LS', ],
             'Linie energetyczne na pozostałych użytkach': ["""SELECT * FROM PISMO_LINIE_ENERG_POZOSTALE WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'L_ENERG_NLS', ],
             'Leśny Materiał Podstawowy': ["""SELECT * FROM PISMO_LMP WHERE [Adres leśny] LIKE ?""", False, 'LMP', ],
             'Bloki upraw pochodnych': ["""SELECT * FROM PISMO_BUP WHERE [Adres leśny] LIKE ?""", False, 'BUP', ],
             'Uprawy pochodne': ["""SELECT * FROM PISMO_UPR_POCH WHERE [Adres leśny] LIKE ?""", False, 'UP', ],
             'Gospodarstwa specjalne': ["""SELECT * FROM PISMO_GOSP_SPEC WHERE [Adres leśny] LIKE ?""", False, 'GOSP_SP', ],
             'Proponowane zmiany Typów Siedliskowych Lasu': ["""SELECT * FROM PISMO_ZMIANA_TSL WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'PROP_ZM_TSL', ],
             'Proponowane lasy ochronne (sum)': ["""SELECT F_AROD_CATEGORY.PROT_CATEGORY_CD, Sum(F_SUBAREA.SUB_AREA) AS SumaOfSUB_AREA FROM (F_ARODES INNER JOIN (F_GROUP_CATEGORY INNER JOIN F_SUBAREA ON F_GROUP_CATEGORY.AREA_TYPE_CD = F_SUBAREA.AREA_TYPE_CD) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_CATEGORY ON F_SUBAREA.ARODES_INT_NUM = F_AROD_CATEGORY.ARODES_INT_NUM WHERE (((F_GROUP_CATEGORY.SUPERGR_CAT_NAME) In ('grunty leśne zalesione','grunty leśne niezalesione')) AND ((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True)) GROUP BY F_AROD_CATEGORY.PROT_CATEGORY_CD;""", False, 'PROP_L_OCHR', ],
             'Użytki ekologiczne': ["""SELECT * FROM PISMO_UE WHERE [Adres leśny] LIKE ?""", False, 'UEK', ],
             'Strefy ochrony': ["""SELECT * FROM PISMO_STREFY WHERE [Adres leśny] LIKE ?""", False, 'STR_OCHR', ],
             'Rośliny objęte monitoringiem': ["""SELECT * FROM PISMO_ROSLINY WHERE [Adres leśny] LIKE ?""", False, 'ROM', ],
             'Pozostałe osobliwości': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST AS [Adres leśny], F_SUBAREA.AREA_TYPE_CD AS [Rodzaj powierzchni], F_AROD_PHENOMENA.PHENOMENA_CD AS [Kod osobliwości], F_TREE_SPECIES.SPECIES_CD, F_AROD_PHENOMENA.NATURE_MON_FL, F_AROD_PHENOMENA.LOCATION_CD AS Lokalizacja, F_SUBAREA.SUBAREA_INFO AS [Informacje dodatkowe] FROM F_TREE_SPECIES RIGHT JOIN ((F_ARODES INNER JOIN ([skrocony_opis-taks] RIGHT JOIN F_SUBAREA ON [skrocony_opis-taks].ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_PHENOMENA ON F_SUBAREA.ARODES_INT_NUM = F_AROD_PHENOMENA.ARODES_INT_NUM) ON F_TREE_SPECIES.SPECIES_CD = F_AROD_PHENOMENA.SPECIES_CD WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_PHENOMENA.PHENOMENA_CD) Not Like 'PŁAT ROŚ') AND ((F_ARODES.TEMP_ACT_ADRESS)=True)) ORDER BY F_ARODES.ORDER_KEY;""", False, 'POZ_OSOB', ],
             'Ekosystemy referencyjne istniejące': ["""SELECT * FROM PISMO_EKO_REF WHERE [Adres leśny] LIKE ?""", False, 'EKO_REF_IST', ],
             'Ekosystemy referencyjne proponowane': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST AS [Adres leśny], F_SUBAREA.AREA_TYPE_CD AS [Rodzaj pow], F_SUBAREA.SUB_AREA AS Pow, F_SUBAREA.SUBAREA_INFO AS [Informacje dodatkowe], F_SUBAREA.FOREST_FUNC_CD AS [Funkcja lasu], [skrocony_opis-taks].ot, F_SUBAREA.SILVICULTURE_CD FROM F_ARODES INNER JOIN ([skrocony_opis-taks] RIGHT JOIN F_SUBAREA ON [skrocony_opis-taks].ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM WHERE (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_SUBAREA.SUBAREA_INFO) Like '%EKO_REF%' And (F_SUBAREA.SUBAREA_INFO) Not Like '%prop.EKO_REF%') AND ((F_ARODES.TEMP_ACT_ADRESS)=True)) ORDER BY F_ARODES.ORDER_KEY;""", False, 'EKO_REF_PROP', ],
             'Siedliska przyrodnicze': ["""SELECT * FROM PISMO_SIEDLISKA_PRZYR WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'SIEDL', ],
             'Halizny, płazowiny': ["""SELECT * FROM PISMO_HALIZNY WHERE [Adres leśny] LIKE ?""", False, 'HAL_PLAZ', ],
             'Sukcesje i do szczególnej ochrony': ["""SELECT * FROM PISMO_SUKCESJA WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'SUKC_DSO', ],
             'Uprawy otwarte': ['SQL3', False, 'UPR_O', ], #achtung
             'Uprawy i młodniki po rębniach złoż.': ['SQL3', False, 'UPR_MLOD_ZLO', ], #achtung
             'Zręby istniejące': ["""SELECT * FROM PISMO_ZREBY WHERE [Adres leśny] LIKE ?""", False, 'ZREB_IST', ],
             'KO, KDO': ["""SELECT * FROM PISMO_KO_KDO WHERE [Adres leśny] LIKE ?""", False, 'KO_KDO', ],
             'Drzewostany z brakiem wskazań': ["""SELECT * FROM PISMO_BRAK_WSKAZAN WHERE [Adres leśny] LIKE ?""", False, 'DST_BRK_WSK', ],
             'Przebudowa A': ["""SELECT * FROM PISMO_PRZEBUD_A WHERE [Adres leśny] LIKE ?""", False, 'PRZEB_A', ],
             'Przebudowa B': ["""SELECT * FROM PISMO_PRZEBUD_B WHERE [Adres leśny] LIKE ?""", False, 'PRZEB_B', ],
             'Przebudowa C': ["""SELECT * FROM PISMO_PRZEBUD_C WHERE [Adres leśny] LIKE ?""", False, 'PRZEB_C', ],
             'Poletka łowieckie': ["""SELECT * FROM PISMO_POLETKA_ŁOW WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'POL_LOW', ],
             'Poletka łowieckie PNSW': ["""SELECT * FROM PISMO_POLETKA_ŁOW_PNSW WHERE [Adres lesny] LIKE ?""", False, 'POL_LOW_PNSW', ],
             'Wprowadzanie podszytów': ["""SELECT * FROM PISMO_WPR_PODSZ WHERE [Adres leśny] LIKE ?""", False, 'WPR_PODSZ', ],
             'Odnawianie luk': ["""SELECT * FROM PISMO_ODN_LUK WHERE [Adres leśny] LIKE ?""", False, 'ODN_LUK', ],
             'Drzewostany brane pod uwagę przy konstrukcji planu cięć': ["""SELECT * FROM PISMO_REBNE WHERE [Adres leśny] LIKE ?""", False, 'DST_BR_KONSTR_CIEC', ],
             'Stanowiska archeologiczne': ["""SELECT * FROM PISMO_ARCHEO WHERE [Adres leśny] LIKE ?""", False, 'ARCHEO', ],
             'Biogrupy w 1 klasie wieku': ["""SELECT * FROM PISMO_BIOGRUPY WHERE TEMP_ADRESS_FOREST LIKE ?""", False, 'BIO_1KW', ],
             'Biogrupy na zrębach': ["""SELECT F_ARODES.TEMP_ADRESS_FOREST, F_AROD_SPECIALAREA.SPECIAL_AREA_CD, F_AROD_SPECIALAREA.LOCATION_CD, F_AROD_SPECIALAREA.SPECIAL_AREA, GAT_PAN.SUBAREA_INFO, F_SPECIES_SPAREA.SPECIES_CD, F_SPECIES_SPAREA.SP_AGE FROM ((F_ARODES INNER JOIN F_AROD_SPECIALAREA ON F_ARODES.ARODES_INT_NUM = F_AROD_SPECIALAREA.ARODES_INT_NUM) INNER JOIN GAT_PAN ON F_AROD_SPECIALAREA.ARODES_INT_NUM = GAT_PAN.ARODES_INT_NUM) INNER JOIN F_SPECIES_SPAREA ON (F_AROD_SPECIALAREA.AROD_SPAREA_ORDER = F_SPECIES_SPAREA.AROD_SPAREA_ORDER) AND (F_AROD_SPECIALAREA.ARODES_INT_NUM = F_SPECIES_SPAREA.ARODES_INT_NUM) AND (F_AROD_SPECIALAREA.SPECIAL_AREA_CD = F_SPECIES_SPAREA.SPECIAL_AREA_CD) GROUP BY F_ARODES.TEMP_ADRESS_FOREST, F_AROD_SPECIALAREA.SPECIAL_AREA_CD, F_AROD_SPECIALAREA.LOCATION_CD, F_AROD_SPECIALAREA.SPECIAL_AREA, GAT_PAN.SUBAREA_INFO, F_SPECIES_SPAREA.SPECIES_CD, F_SPECIES_SPAREA.SP_AGE, F_ARODES.ARODES_INT_NUM, F_ARODES.TEMP_ACT_ADRESS, GAT_PAN.SPECIES_AGE, F_SPECIES_SPAREA.SP_RANK_ORDER HAVING (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_SPECIALAREA.SPECIAL_AREA_CD)='KĘPA') AND ((F_ARODES.TEMP_ACT_ADRESS)=True) AND ((GAT_PAN.SPECIES_AGE)<21) AND ((F_SPECIES_SPAREA.SP_RANK_ORDER)=1));""", False, 'BIO_ZR', ],
             'Rezerwaty przyrody': ["""SELECT F_ARODES.[TEMP_ADRESS_FOREST], F_SUBAREA.[AREA_TYPE_CD], F_SUBAREA.[SUB_AREA], F_SUBAREA.[SUBAREA_INFO], F_SUBAREA.[FOREST_FUNC_CD] FROM F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.[ARODES_INT_NUM] = F_SUBAREA.[ARODES_INT_NUM] GROUP BY F_ARODES.[ARODES_INT_NUM], F_ARODES.[TEMP_ADRESS_FOREST], F_ARODES.[TEMP_ACT_ADRESS], F_SUBAREA.[AREA_TYPE_CD], F_SUBAREA.[SUB_AREA], F_SUBAREA.[SUBAREA_INFO], F_SUBAREA.[FOREST_FUNC_CD] HAVING (((F_ARODES.[TEMP_ADRESS_FOREST]) Like ?) AND ((F_ARODES.[TEMP_ACT_ADRESS])=True) AND ((F_SUBAREA.[FOREST_FUNC_CD])='REZ'));""", False, 'REZ_PRZ', ]

             },

    }
    return RAPORTS
