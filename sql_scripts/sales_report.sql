-- Отчёт по продажам
-- Параметры: :date_from, :date_to
SELECT
    s.id                    AS "№ Заказа",
    s.created_at::date      AS "Дата",
    c.name                  AS "Клиент",
    p.name                  AS "Товар",
    sd.quantity             AS "Кол-во",
    sd.price                AS "Цена",
    (sd.quantity * sd.price) AS "Сумма"
FROM sales s
JOIN clients   c  ON c.id = s.client_id
JOIN sale_details sd ON sd.sale_id = s.id
JOIN products  p  ON p.id = sd.product_id
WHERE s.created_at::date BETWEEN :date_from AND :date_to
ORDER BY s.created_at DESC;
