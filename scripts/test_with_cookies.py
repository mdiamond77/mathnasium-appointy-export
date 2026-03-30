"""Quick test: make the GraphQL export call using existing browser cookies."""
import os
from datetime import date, timedelta
from scripts.appointy_export import (
    get_date_range, to_utc_iso_start, to_utc_iso_end,
    COMPANY_ID, GROUP_ID, LOCATION_ID, USER_ID, BASE_URL,
    GRAPHQL_QUERY, EXPORTED_FIELDS,
)
import requests

cf_clearance    = os.environ["COOKIE_CF"]
ory_session     = os.environ["COOKIE_SESSION"]
oidc            = os.environ["COOKIE_OIDC"]

today = date.today()
start, end = get_date_range(today)

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Content-Type": "application/json",
    "Origin": BASE_URL,
    "Referer": f"{BASE_URL}/",
})
id_gen = os.environ.get("COOKIE_IDGEN", "")

# Send cookies as a raw header string — exactly like a browser would
cookie_parts = [
    f"cf_clearance={cf_clearance}",
    f"ory_hydra_session={ory_session}",
    f"oidc={oidc}",
    f"partitioned-cookie={oidc}",
]
if id_gen:
    cookie_parts.append(f"idGenSession={id_gen}")
session.headers["Cookie"] = "; ".join(cookie_parts)

print(f"Testing export: {start} → {end}")
resp = session.post(
    f"{BASE_URL}/graphql",
    params={
        "companyId":  COMPANY_ID,
        "groupId":    GROUP_ID,
        "locationId": LOCATION_ID,
        "queryId":    "CompanyAppointmentReportQuery",
        "userId":     USER_ID,
    },
    json={
        "id":    "CompanyAppointmentReportQuery",
        "query": GRAPHQL_QUERY,
        "variables": {
            "accessContact": True,
            "additionalReportsFilters": None,
            "appointmentDate": {
                "appointmentDate": {
                    "startTime": to_utc_iso_start(start),
                    "endTime":   to_utc_iso_end(end),
                }
            },
            "bookingDate": None, "consumerEmail": None, "consumerName": None,
            "consumerTag": None, "customerEmail": None, "customerName": None,
            "displayCustomerId": None, "dropDownFilters": None,
            "employeeEmail": None, "employeeName": None,
            "export": True,
            "exportedFields": EXPORTED_FIELDS,
            "hasExtendedFormFields": False,
            "limit": 200, "locationIds": [], "mathnasium": True,
            "offset": 0, "orderBy": None, "packageCode": None,
            "parent": f"{GROUP_ID}/{COMPANY_ID}",
            "paymentMethod": None, "price": None, "serviceTitle": None,
            "skipBuyerOrganization": True, "sortBy": None, "statusFilter": None,
        },
    },
)
print(f"Status: {resp.status_code}")
print(f"Response: {resp.text[:300]}")
