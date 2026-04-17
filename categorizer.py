"""
Transaction categorization engine for nonprofit organizations.
Maps bank statement descriptions to nonprofit-specific revenue and expense categories.
"""

import re
from typing import Optional

REVENUE_CATEGORIES = [
    "Donations & Contributions",
    "Donations & Contributions: Tournament Sponsor",
    "Donations & Contributions: Employer Match",
    "Grants",
    "Program Service Revenue",
    "Program Service Revenue: Dirty Bird Tournament",
    "Program Service Revenue: Skills & Drills",
    "Investment Income",
    "Fundraising Event Revenue",
    "Other Revenue",
]

EXPENSE_CATEGORIES = [
    "Program Services",
    "Program Services: Dirty Birds Tournament",
    "Program Services: Travel Tournament",
    "Program Services: Skills & Drills",
    "Management & General",
    "Fundraising",
    "Facilities & Occupancy",
    "Salaries & Benefits",
    "Supplies & Materials",
    "League Equipment",
    "Travel & Transportation",
    "Professional Services",
    "Insurance",
    "Scholarship",
    "Internal Account Transfer",
    "Other Expenses",
]

FUNCTIONAL_CATEGORIES = ["Program Services", "Management & General", "Fundraising"]

ALL_CATEGORIES = REVENUE_CATEGORIES + EXPENSE_CATEGORIES

SUBCATEGORY_MAP = {
    "Donations & Contributions: Tournament Sponsor": "Donations & Contributions",
    "Donations & Contributions: Employer Match": "Donations & Contributions",
    "Program Service Revenue: Dirty Bird Tournament": "Program Service Revenue",
    "Program Service Revenue: Skills & Drills": "Program Service Revenue",
    "Program Services: Dirty Birds Tournament": "Program Services",
    "Program Services: Travel Tournament": "Program Services",
    "Program Services: Skills & Drills": "Program Services",
}


def get_parent_category(category: str) -> str:
    """Return the parent category, or the category itself if it has no parent."""
    return SUBCATEGORY_MAP.get(category, category)

KEYWORD_MAP = {
    "Donations & Contributions": [
        r"donat", r"contribut", r"tithe", r"offering", r"gift",
        r"pledge", r"charit", r"philanthropi", r"benevo",
        r"mobile\s+deposit", r"descriptive\s+deposit",
    ],
    "Grants": [
        r"grant", r"award", r"foundation\s+(?:fund|pay)", r"government\s+fund",
        r"federal\s+fund", r"state\s+fund",
    ],
    "Program Service Revenue": [
        r"program\s+fee", r"tuition", r"registration", r"enrollment",
        r"service\s+fee", r"membership\s+due", r"membership\s+fee",
        r"class\s+fee", r"workshop\s+fee", r"training\s+fee",
        r"venmo\s+.*cashout", r"cashout\s+.*deposit",
        r"zelle.*(?:payment|deposit)", r"paypal.*(?:transfer|deposit)",
    ],
    "Investment Income": [
        r"interest\s+(?:income|earned|payment)", r"dividend", r"investment",
        r"capital\s+gain", r"yield", r"interest\s+from",
    ],
    "Fundraising Event Revenue": [
        r"fundrais", r"gala", r"auction", r"benefit\s+dinner",
        r"raffle", r"ticket\s+sale", r"event\s+revenue",
    ],
    "Salaries & Benefits": [
        r"payroll", r"salary", r"salaries", r"wage", r"compensation",
        r"benefit", r"health\s+ins", r"dental", r"retirement",
        r"401k", r"pension", r"worker.?s?\s+comp", r"adp", r"gusto",
        r"paychex", r"direct\s+deposit",
    ],
    "Facilities & Occupancy": [
        r"rent", r"lease", r"mortgage", r"utilit", r"electric",
        r"gas\s+bill", r"water\s+bill", r"internet", r"phone\s+bill",
        r"maintenance", r"repair", r"janitorial", r"cleaning",
        r"property\s+tax", r"landscap",
        r"ice\s+(?:rink|arena|sport|center)", r"sno\s+king", r"ice\s+sport",
        r"skating\s+(?:rink|center|arena)", r"field\s+(?:rental|reserv)",
        r"gym\s+(?:rental|fee)", r"court\s+(?:rental|fee)",
        r"venue\s+(?:rental|fee)",
    ],
    "Supplies & Materials": [
        r"supply", r"supplies", r"material", r"office\s+depot",
        r"staples", r"amazon", r"walmart", r"target",
        r"paper", r"ink", r"toner",
        r"dollar\s*tree", r"fedex\s+offic", r"fedex\s+office",
        r"ups\s+store", r"kinkos",
    ],
    "League Equipment": [
        r"hockeytron", r"hockey\s*monkey", r"pure\s*hockey",
        r"dick.?s\s+sport", r"sport\s*check",
        r"equipment", r"gear", r"jersey", r"jerseys",
        r"uniform", r"puck", r"stick", r"skate",
        r"helmet", r"glove", r"pad\b", r"pads\b",
        r"protective", r"net\b", r"goal\s+equip",
    ],
    "Travel & Transportation": [
        r"travel", r"flight", r"airfare", r"airline", r"hotel",
        r"lodging", r"uber", r"lyft", r"taxi", r"mileage",
        r"gas\s+station", r"fuel", r"parking", r"toll",
    ],
    "Professional Services": [
        r"legal", r"attorney", r"lawyer", r"accounting", r"audit",
        r"consult", r"contractor", r"freelanc", r"design\s+serv",
        r"marketing\s+serv", r"advisory", r"cpa",
    ],
    "Insurance": [
        r"insurance", r"liability", r"coverage", r"premium",
        r"policy\s+(?:pay|renew)", r"indemnity",
    ],
    "Fundraising": [
        r"fundrais(?:ing)?\s+(?:expense|cost|supply)", r"donor\s+manag",
        r"campaign", r"solicitation", r"appeal\s+mail",
        r"event\s+(?:cost|expense|decor|cater)",
    ],
    "Management & General": [
        r"admin", r"management\s+fee", r"bank\s+(?:fee|charge|service)",
        r"license", r"permit", r"filing\s+fee", r"registration\s+fee",
        r"software\s+(?:sub|license)", r"quickbooks", r"xero",
        r"subscription", r"dues", r"postage", r"shipping",
        r"google\s*\*", r"google\s+one", r"microsoft", r"adobe",
        r"intuit", r"square\s+(?:fee|service)",
        r"monthly\s+(?:fee|charge|service)",
    ],
    "Program Services": [
        r"program\s+(?:expense|cost|supply|material)",
        r"curriculum", r"workshop\s+(?:expense|cost)",
        r"participant", r"client\s+serv", r"outreach",
        r"community\s+serv", r"mission",
        r"league\s+(?:fee|due|registr)",
        r"(?:team|player)\s+(?:fee|registr)", r"referee",
        r"winterhawk", r"cascade\s+athletic",
        r"seattle\s+women", r"spo\*seattle",
        r"usa\s+hockey", r"game\s+(?:fee|official|ice)",
    ],
    "Program Services: Dirty Birds Tournament": [
        r"dirty\s*bird", r"dirty\s*birds\s*(?:tourn|event|fee)",
    ],
    "Program Services: Travel Tournament": [
        r"travel\s+tourn", r"away\s+tourn", r"road\s+tourn",
        r"tournament\s+travel", r"tournament\s+entry",
        r"tournament\s+(?:fee|registr)",
        r"tournament",
    ],
    "Program Services: Skills & Drills": [
        r"skills?\s*(?:&|and)\s*drills?",
        r"skills?\s+clinic", r"drills?\s+session",
        r"skills?\s+training", r"practice\s+session",
    ],
    "Program Service Revenue: Dirty Bird Tournament": [
        r"dirty\s*bird", r"dirty\s*birds?\s*(?:tourn|registr|entry|fee)",
    ],
    "Program Service Revenue: Skills & Drills": [
        r"skills?\s*(?:&|and)\s*drills?",
        r"skills?\s+(?:clinic|class|session)\s+(?:fee|revenue|income)",
    ],
    "Donations & Contributions: Tournament Sponsor": [
        r"(?:tourn|event)\s*sponsor", r"sponsor.*(?:tourn|event)",
        r"sponsorship",
    ],
    "Donations & Contributions: Employer Match": [
        r"employer\s+match", r"company\s+match", r"corporate\s+match",
        r"matching\s+(?:gift|donat|contribut|fund)",
        r"benevity", r"yourcause", r"double\s+the\s+donat",
    ],
    "Scholarship": [
        r"scholar", r"financial\s+aid", r"player\s+assist",
        r"tuition\s+assist", r"need.based\s+(?:aid|assist)",
    ],
}


TRANSFER_PATTERNS = [
    r"transfer\s+(?:from|to)", r"internet\s+transfer",
    r"new\s+account\s+deposit", r"funds\s+transfe",
    r"acctverify", r"verifybank",
]


def categorize_transaction(description: str, amount: float) -> str:
    """Categorize a transaction based on its description and amount."""
    if not description:
        return "Other Revenue" if amount > 0 else "Other Expenses"

    desc_lower = description.lower().strip()

    for pattern in TRANSFER_PATTERNS:
        if re.search(pattern, desc_lower):
            return "Internal Account Transfer"

    if amount <= 0 and re.search(r"venmo\s+.*payment|payment.*venmo", desc_lower):
        return "Program Services"

    for category, patterns in KEYWORD_MAP.items():
        for pattern in patterns:
            if re.search(pattern, desc_lower):
                is_revenue_cat = category in REVENUE_CATEGORIES
                if amount > 0 and is_revenue_cat:
                    return category
                if amount <= 0 and not is_revenue_cat:
                    return category
                if amount > 0 and not is_revenue_cat:
                    return "Other Revenue"
                if amount <= 0 and is_revenue_cat:
                    return "Other Expenses"

    return "Other Revenue" if amount > 0 else "Other Expenses"


def get_functional_classification(category: str) -> str:
    """Map a detailed expense category to a functional classification."""
    if category in REVENUE_CATEGORIES:
        return "N/A"

    functional_map = {
        "Program Services": "Program Services",
        "Program Services: Dirty Birds Tournament": "Program Services",
        "Program Services: Travel Tournament": "Program Services",
        "Program Services: Skills & Drills": "Program Services",
        "Salaries & Benefits": "Program Services",
        "Supplies & Materials": "Program Services",
        "League Equipment": "Program Services",
        "Travel & Transportation": "Program Services",
        "Scholarship": "Program Services",
        "Management & General": "Management & General",
        "Professional Services": "Management & General",
        "Insurance": "Management & General",
        "Facilities & Occupancy": "Management & General",
        "Internal Account Transfer": "N/A",
        "Fundraising": "Fundraising",
        "Other Expenses": "Management & General",
    }
    return functional_map.get(category, "Management & General")


def get_category_type(category: str) -> str:
    """Return 'Revenue', 'Expense', or 'Transfer' for a given category."""
    if category == "Internal Account Transfer":
        return "Transfer"
    if category in REVENUE_CATEGORIES:
        return "Revenue"
    return "Expense"
