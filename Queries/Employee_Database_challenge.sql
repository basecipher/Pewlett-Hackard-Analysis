SELECT 
	e.emp_no,
	e.first_name,
	e.last_name,
	e.gender,
	t.title,
	t.to_date
INTO retirement_titles
FROM employees as e
INNER JOIN titles as t
ON (e.emp_no = t.emp_no)
WHERE (e.birth_date BETWEEN '1952-01-01' AND '1955-12-31')
	AND (e.hire_date BETWEEN '1985-01-01' AND '1988-12-31')
	AND (t.to_date = '9999-01-01')
ORDER BY e.emp_no;



-- Use Dictinct with Orderby to remove duplicate rows
SELECT DISTINCT ON emp_no, first_name, last_name, title,
INTO unique_titles
FROM retirement_titles
ORDER BY (t.to_date = '9999-01-01') DESC;