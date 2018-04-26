get_bank_info = '''SELECT
  b.NAMEP AS CONTRACT_BANK_NAME,
  CONCAT_WS(', ', b.NNP, b.ADR) AS CONTRACT_BANK_ADDRESS
  FROM bnkseek b
  WHERE b.NEWNUM = '%(CONTRACT_BIC)s'
  ;'''

get_procedure_id = '''SELECT
  p.id
FROM procedures AS p
WHERE p.registrationNumber = '%(PROCEDURE_NUMBER_CH)s'
AND p.archive = 0;'''

get_lot_customer_id = '''SELECT
  lc.id
FROM procedures AS p
  JOIN lot AS l
    ON l.procedureId = p.id
    AND l.archive = 0
  JOIN lotCustomer AS lc
    ON lc.lotId = l.id
    AND lc.archive = 0
WHERE p.registrationNumber = '%(PROCEDURE_NUMBER_CH)s'
AND p.archive = 0'''

get_last_insert_id = '''SELECT LAST_INSERT_ID()'''


procedure_guarantee_update = '''UPDATE procedures AS p
SET p.guaranteeFee = '5500.00',
    p.guaranteeFeePercent = NULL
WHERE p.id = '%(PROCEDURE_ID)s'
;'''

request_provision_insert = '''INSERT INTO provision
SET
  percent = '%(REQUEST_PROVISION_PERCENT)s',
  receiver = 'operator',
  receiverText = 'Р/С:40702810700030004213, БИК:044525187, К/С:30101810700000000187, БАНК: БАНК ВТБ (ПАО) Г МОСКВА',
  `value` = '%(REQUEST_PROVISION_RUB)s',
  discriminator = 'request'
  ;'''

set_request_provision_id_update = '''UPDATE lotCustomer AS lc
SET lc.requestProvisionId = '%(REQUEST_PROVISION_ID)s'
WHERE lc.id = '%(LOT_CUSTOMER_ID)s'
;'''

provision_insert = '''INSERT INTO provision
  SET
  `order` = 'Согласно документации',
  percent = '%(CONTRACT_PROVISION_PERCENT)s',
  receiverText = 'Р/С:%(CONTRACT_RASCHET_SCHET)s, Л/С %(CONTRACT_LITS_SCHET)s, БИК:%(CONTRACT_BIC)s, %(CONTRACT_BANK_NAME)s',
  `value` = '%(CONTRACT_PROVISION_RUB)s',
  fullName = '%(CONTRACT_ORGANISATION_NAME)s',
  personalAccount = '%(CONTRACT_LITS_SCHET)s',
  paymentAccount = '%(CONTRACT_RASCHET_SCHET)s',
  bic = '%(CONTRACT_BIC)s',
  bankName = '%(CONTRACT_BANK_NAME)s',
  bankAddress = '%(CONTRACT_BANK_ADDRESS)s',
  discriminator = 'contract'
;'''

contract_provision_id_update = '''UPDATE lotCustomer AS lc
SET lc.contractProvisionId = '%(CONTRACT_PROVISION_ID)s'
WHERE lc.id = '%(LOT_CUSTOMER_ID)s'
;'''

lot_customer_view_update = '''UPDATE lotCustomer AS lc
SET `enabledRequestProvision` = %(enabledRequestProvision)s,
    `enabledContractProvision` = %(enabledContractProvision)s
WHERE lc.id = '%(LOT_CUSTOMER_ID)s'
;'''









