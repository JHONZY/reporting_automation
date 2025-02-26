SELECT
    leads.`leads_chcode` AS 'CH CODE',
    leads.`leads_acctno` AS 'PN NUMBER',
    leads_user.`users_username` AS 'AGENT TAG',
    leads.`leads_chname` AS 'NAME',
    leads.`leads_email` AS 'EMAIL',
    leads.`leads_mobile_phone` AS 'MOBILE',
    leads.`leads_ob` AS 'OUTSTANDING BALANCE',
    leads.`leads_prin` AS 'ENDORSE OB',
    leads.`leads_full_address` AS 'ADDRESS',
    leads.`leads_full_saddress` AS 'SECONDARY',
    DATE_FORMAT(leads.`leads_endo_date`, '%m/%d/%Y') AS 'ENDO DATE',
    DATE_FORMAT(leads.`leads_cutoff`, '%m/%d/%Y') AS 'PULL OUT DATE',
    dv1.`dynamic_value_name` AS 'DPD',
    CASE
        WHEN leads_user.`users_username` = 'MSPM' THEN 
            (SELECT CASE 
                WHEN leads_result.`leads_result_users` = 5505 THEN 'CHARINA ALBAR'
                WHEN leads_result.`leads_result_users` = 5581 THEN 'BABY JANE MENDOZA'
                WHEN leads_result.`leads_result_users` = 6463 THEN 'HANNAH DALUDADO' 
            END
            FROM bcrm.leads_result
            WHERE leads_result.leads_result_lead = leads.leads_id
            ORDER BY leads_result.leads_result_barcode_date DESC
            LIMIT 1
            )
        WHEN leads_user.`users_username` = 'PCAA' THEN 'CHARINA ALBAR'
        WHEN leads_user.`users_username` = 'MBSM' THEN 'BABY JANE MENDOZA'
        WHEN leads_user.`users_username` = 'PHED' THEN 'HANNAH DALUDADO' 
        ELSE leads_user.`users_username`
    END AS `LastTouch`,
    DATE_FORMAT(leads.`leads_ts`, '%m/%d/%Y %h:%i %p') AS 'DateProcessed'
FROM `bcrm`.`leads`
LEFT JOIN `bcrm`.`dynamic_value` AS dv1 ON dv1.`dynamic_value_lead_id` = leads.`leads_id`
LEFT JOIN `bcrm`.`users` AS leads_user ON leads_user.`users_id` = leads.`leads_users_id`
WHERE leads.`leads_client_id` = 191
AND leads.`leads_users_id` <> 659
GROUP BY leads.`leads_chcode`
ORDER BY `leads_endo_date` DESC;
