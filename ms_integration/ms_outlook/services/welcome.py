# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Temporary passwords and the welcome email.

A new random password per person, never stored: it is set in Microsoft 365
(change required at first sign-in) and emailed once to the employee's
personal address.
"""

import secrets
import string

import frappe
from frappe import _

from ms_integration.ms_core.services.settings import get_connection
from ms_integration.ms_outlook.services.account_api import send_mail

LOGIN_URL = "https://www.office.com"
SYMBOLS = "!@#$%*?-"
DEFAULT_SUBJECT = "Your Microsoft 365 account"
DEFAULT_MESSAGE = """<p>Hello {{ doc.first_name or doc.employee_name }},</p>
<p>Your Microsoft 365 account is ready. Sign in at
<a href="{{ login_url }}">{{ login_url }}</a> with:</p>
<p><b>Email:</b> {{ email }}<br><b>Temporary password:</b> {{ password }}</p>
<p>You will be asked to choose your own password the first time you sign in.</p>"""


def new_password(length: int = 16) -> str:
	"""
	Random password meeting Microsoft's complexity rules.

	Parameters:
	        length (int, optional): Length (default 16).

	Returns:
	        str: Password with upper, lower, digit and symbol.
	"""
	pools = [string.ascii_uppercase, string.ascii_lowercase, string.digits, SYMBOLS]
	chars = [secrets.choice(pool) for pool in pools]
	everything = "".join(pools)
	chars += [secrets.choice(everything) for _ in range(length - len(chars))]
	secrets.SystemRandom().shuffle(chars)
	return "".join(chars)


def send_welcome(client, employee, email: str, password: str, settings) -> str:
	"""
	Email the sign-in details to the employee.

	Parameters:
	        client (GraphClient, required): Graph client.
	        employee (Document, required): Employee.
	        email (str, required): New sign-in address.
	        password (str, required): Temporary password.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        str: Recipient address.

	Raises:
	        frappe.ValidationError: When the employee has no address to send to.
	"""
	field = settings.welcome_recipient_field or "personal_email"
	recipient = (employee.get(field) or "").strip()
	if not recipient:
		frappe.throw(
			_("Employee {0} has no {1} to send the welcome email to.").format(employee.name, field)
		)
	subject, html = render(employee, email, password, settings)
	if settings.welcome_via == "ERPNext Email":
		send_through_erpnext(recipient, subject, html)
	else:
		sender = settings.sender_mailbox or (
			f"it.admin@{settings.email_domain}" if get_connection().test_mode else ""
		)
		send_mail(client, sender, recipient, subject, html)
	return recipient


def render(employee, email: str, password: str, settings) -> tuple[str, str]:
	"""
	Subject and HTML from the Email Template, or the built-in message.

	Parameters:
	        employee (Document, required): Employee.
	        email (str, required): Sign-in address.
	        password (str, required): Temporary password.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        tuple[str, str]: Subject and HTML body.
	"""
	context = {"doc": employee, "email": email, "password": password, "login_url": LOGIN_URL}
	subject, body = DEFAULT_SUBJECT, DEFAULT_MESSAGE
	if settings.welcome_email_template:
		template = frappe.get_cached_doc("Email Template", settings.welcome_email_template)
		subject = template.subject or subject
		body = (template.response_html if template.use_html else template.response) or body
	# Template comes from Microsoft 365 Settings (System Manager only) or the built-in text.
	subject = frappe.render_template(subject, context)  # nosemgrep
	return subject, frappe.render_template(body, context)  # nosemgrep


def send_through_erpnext(recipient: str, subject: str, html: str) -> None:
	"""
	Send now through the site's email account and delete the queued copy,
	so the password does not stay in the Email Queue.

	Parameters:
	        recipient (str, required): Address.
	        subject (str, required): Subject.
	        html (str, required): HTML body.

	Returns:
	        None
	"""
	queue = frappe.sendmail(recipients=[recipient], subject=subject, message=html, now=True)
	queue = queue[0] if isinstance(queue, list) and queue else queue
	if not queue:
		frappe.throw(_("The welcome email could not be queued."))
	status = frappe.db.get_value("Email Queue", queue.name, "status")
	frappe.delete_doc("Email Queue", queue.name, ignore_permissions=True, force=True)
	if status != "Sent":
		frappe.throw(_("The welcome email could not be sent (status: {0}).").format(status))
