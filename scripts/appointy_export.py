"""
Appointy Nightly Export Script
Logs into Appointy via HTTP and triggers the Appointment Detailed Report
email export for Sunday of the current week through 10 weeks out.
"""

import os
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import requests

# ── Location constants (specific to this Mathnasium franchise) ────────────────
COMPANY_ID  = "com_01G6J3TV8M70JXM8M4972NA39D"
GROUP_ID    = "grp_01F1QGKK7M8ZY6D72JD7KC0K00"
LOCATION_ID = "loc_01G6J4T8K9GZYYKM2EP6AQP175"
USER_ID     = "usr_01GAJE0K0MTNK147VKW6J0AQ4S"
BASE_URL    = "https://mathnasium-admin.appointy.com"
EASTERN     = ZoneInfo("America/New_York")

EXPORTED_FIELDS = [
    "location", "appointmentDate", "consumers", "customerName",
    "customerEmail", "address", "serviceTitle", "employeeName",
    "duration", "customStatus", "status",
]

GRAPHQL_QUERY = """query CompanyAppointmentReportQuery(
  $accessContact: Boolean
  $parent: String
  $locationIds: [String]
  $limit: Int
  $offset: Int
  $appointmentDate: getCompanyAppointmentReportRequestAppointmentDate
  $bookingDate: getCompanyAppointmentReportRequestBookingDate
  $consumerEmail: String
  $consumerTag: String
  $consumerName: String
  $employeeEmail: String
  $employeeName: String
  $price: PriceInput
  $serviceTitle: String
  $statusFilter: AppointmentStatusFilterInput
  $export: Boolean
  $mathnasium: Boolean!
  $customerEmail: String
  $customerName: String
  $exportedFields: [String]
  $hasExtendedFormFields: Boolean!
  $dropDownFilters: ReportsDropDownFilterInput
  $packageCode: String
  $paymentMethod: PaymentMethod
  $sortBy: AppointmentReportSortBy
  $orderBy: OrderBy
  $displayCustomerId: String
  $additionalReportsFilters: AdditionalReportsFilterInput
  $skipBuyerOrganization: Boolean
) {
  companyAppointmentReport(accessContact: $accessContact, parent: $parent, locationIds: $locationIds, export: $export, limit: $limit, offset: $offset, appointmentDate: $appointmentDate, bookingDate: $bookingDate, consumerEmail: $consumerEmail, consumerTag: $consumerTag, consumerName: $consumerName, employeeEmail: $employeeEmail, employeeName: $employeeName, price: $price, serviceTitle: $serviceTitle, statusFilter: $statusFilter, customerEmail: $customerEmail, customerName: $customerName, exportedFields: $exportedFields, dropDownFilters: $dropDownFilters, packageCode: $packageCode, paymentMethod: $paymentMethod, sortBy: $sortBy, orderBy: $orderBy, displayCustomerId: $displayCustomerId, additionalReportsFilters: $additionalReportsFilters) {
    edges {
      node {
        data {
          appointmentId
          quantity
          appointmentDate {
            startTime
            endTime
          }
          paymentStatus @skip(if: $mathnasium)
          bookingDate
          bookedBy {
            email
            firstName
            lastName
            userId
          }
          employeeExtendedFormFieldValues @include(if: $hasExtendedFormFields) {
            addressKey {
              country
              latitude
              locality
              longitude
              postalCode
              region
              streetAddress
            }
            boolKey
            floatKey
            intArrKey
            intKey
            key
            stringArrKey
            stringKey
          }
          employeeExtendedFormFields @include(if: $hasExtendedFormFields) {
            formFieldType
            key
            label
            metadata
            multiple
            number
            options {
              name
              value
            }
            placeholder
          }
          extendedFormFields @include(if: $hasExtendedFormFields) {
            formFieldType
            key
            label
            metadata
            multiple
            number
            options {
              name
              value
            }
            placeholder
          }
          extendedFormFieldValues @include(if: $hasExtendedFormFields) {
            addressKey {
              country
              latitude
              locality
              longitude
              postalCode
              region
              streetAddress
            }
            boolKey
            floatKey
            intArrKey
            intKey
            key
            stringArrKey
            stringKey
          }
          employeeSsoId
          consumers {
            tags
            companyId
            extendedFormFieldValues @include(if: $hasExtendedFormFields) {
              addressKey {
                country
                latitude
                locality
                longitude
                postalCode
                region
                streetAddress
              }
              boolKey
              floatKey
              intArrKey
              intKey
              key
              stringArrKey
              stringKey
            }
            extendedFormFields @include(if: $hasExtendedFormFields) {
              formFieldType
              key
              label
              metadata
              multiple
              number
              options {
                name
                value
              }
              placeholder
            }
            details {
              cancelationReason
              consumerGuests
              email
              _id
              customStatus {
                customStatusId
                customStatusName
              }
              consumerData {
                data {
                  __typename
                  ... on ConsumerData_CustomerData {
                    data: customerData {
                      firstName
                      lastName
                    }
                  }
                  ... on ConsumerData_StudentData {
                    data: studentData {
                      firstName
                      lastName
                      metadata @include(if: $mathnasium)
                    }
                  }
                }
              }
              address {
                streetAddress
                country
                region
                locality
                postalCode
              }
            }
            ssoId
            displayCustomerId
          }
          discountCode
          membershipCode
          duration
          employeeEmail
          employeeFirstName
          employeeLastName
          totalPrice {
            amount
            currency
          }
          additionalAmountNote
          buyerOrganization @skip(if: $skipBuyerOrganization) {
            id
            name
          }
          totalDiscountedPrice {
            amount
            currency
          }
          additionalDiscountNote
          isBookedByAdmin
          notes
          resourceTypes {
            title
            capacity
          }
          resources {
            title
            capacity
          }
          serviceTitle
          serviceId
          status
          source
          medium
          campaign
          forms {
            name
            blocks {
              key
              label
            }
          }
          submissionValues {
            key
            response {
              boolean
              decimal
              text
              number
              timestamp
              valueType
            }
          }
          nameOfCancelledBy
          nameOfRescheduledBy
          reason
          serviceCategory
          bookingUrl
        }
        location {
          locationId
          name
          customLocationId
        }
      }
      cursor
    }
    pageInfo {
      previousOffset
      nextOffset
      hasNextPage
      hasPreviousPage
    }
    downloadLimit
    total
    duration
    totalPrice {
      amount
      currency
    }
  }
}
"""


# ── Date helpers ──────────────────────────────────────────────────────────────

def get_week_start(today: date) -> date:
    """Return the most recent Sunday on or before today."""
    days_since_sunday = (today.weekday() + 1) % 7
    return today - timedelta(days=days_since_sunday)


def get_date_range(today: date) -> tuple[date, date]:
    """Return (start, end): Sunday of current week → start + 70 days."""
    start = get_week_start(today)
    end = start + timedelta(days=70)
    return start, end


def to_utc_iso_start(d: date) -> str:
    """Midnight Eastern on date d, expressed as UTC ISO string."""
    dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=EASTERN)
    utc = dt.astimezone(ZoneInfo("UTC"))
    return utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")


def to_utc_iso_end(d: date) -> str:
    """11:59 PM Eastern on date d, expressed as UTC ISO string."""
    dt = datetime(d.year, d.month, d.day, 23, 59, 0, tzinfo=EASTERN)
    utc = dt.astimezone(ZoneInfo("UTC"))
    return utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── Auth ──────────────────────────────────────────────────────────────────────

def login(session: requests.Session, cookie_string: str) -> None:
    """
    Authenticate by injecting the browser cookie string directly.
    The cookie string is copied from Chrome DevTools and stored as a secret.
    """
    session.headers["Cookie"] = cookie_string


# ── Export ────────────────────────────────────────────────────────────────────

def trigger_export(session: requests.Session, start: date, end: date) -> dict:
    """POST the GraphQL export request with export=True to trigger the email."""
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
                "accessContact":           True,
                "additionalReportsFilters": None,
                "appointmentDate": {
                    "appointmentDate": {
                        "startTime": to_utc_iso_start(start),
                        "endTime":   to_utc_iso_end(end),
                    }
                },
                "bookingDate":       None,
                "consumerEmail":     None,
                "consumerName":      None,
                "consumerTag":       None,
                "customerEmail":     None,
                "customerName":      None,
                "displayCustomerId": None,
                "dropDownFilters":   None,
                "employeeEmail":     None,
                "employeeName":      None,
                "export":            True,
                "exportedFields":    EXPORTED_FIELDS,
                "hasExtendedFormFields": False,
                "limit":             200,
                "locationIds":       [],
                "mathnasium":        True,
                "offset":            0,
                "orderBy":           None,
                "packageCode":       None,
                "parent":            f"{GROUP_ID}/{COMPANY_ID}",
                "paymentMethod":     None,
                "price":             None,
                "serviceTitle":      None,
                "skipBuyerOrganization": True,
                "sortBy":            None,
                "statusFilter":      None,
            },
        },
        headers={
            "Content-Type": "application/json",
            "Origin":       BASE_URL,
            "Referer":      f"{BASE_URL}/",
        },
    )
    resp.raise_for_status()
    return resp.json()


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    cookie_string = os.environ["APPOINTY_COOKIES"]

    today = date.today()
    start, end = get_date_range(today)

    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/146.0.0.0 Safari/537.36"
        )
    })

    login(session, cookie_string)
    print(f"Triggering export: {start} → {end}")
    result = trigger_export(session, start, end)

    errors = result.get("errors")
    if errors:
        raise RuntimeError(f"GraphQL error: {errors}")

    print("Export triggered successfully — email should arrive shortly.")


if __name__ == "__main__":
    main()
