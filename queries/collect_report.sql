SELECT
    users.users_username AS Agent,
    leads.leads_chname AS Account_Name,
    leads.leads_acctno AS Account_No,
    leads_status.leads_status_name AS Status,
    leads_substatus.leads_substatus_name AS Substatus,
    leads_result.leads_result_sdate AS Start_Date,
    leads_result.leads_result_edate AS End_Date,
    leads_result.leads_result_amount AS Amount,
    REPLACE(REPLACE(leads_result.leads_result_comment, '1_', ''), '0_', '') AS Notes,
    DATE_FORMAT(leads_result.leads_result_barcode_date, '%m/%d/%y %h:%i %p') AS Barcoded_Date,
    leads.leads_ob AS OB
FROM bcrm.leads_result
LEFT JOIN bcrm.leads ON leads_result.leads_result_lead = leads.leads_id
LEFT JOIN bcrm.leads_status ON leads_result.leads_result_status_id = leads_status.leads_status_id
LEFT JOIN bcrm.leads_substatus ON leads_result.leads_result_substatus_id = leads_substatus.leads_substatus_id
INNER JOIN bcrm.users ON leads_result.leads_result_users = users.users_id
WHERE
    leads_client_id = 146
    AND leads_status.leads_status_name NOT IN ('RETURNS', 'DEAD ACCOUNT')
    AND leads_result.leads_result_barcode_date BETWEEN '2024-11-01' AND '2024-11-29'
    AND leads_result.leads_result_hidden <> 1
ORDER BY leads_result.leads_result_id DESC;
