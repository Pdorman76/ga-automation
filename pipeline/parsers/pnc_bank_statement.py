"""
PNC Bank Statement Parser

This parser extracts data from PNC bank statement PDFs including Bank of America,
KeyBank, and PNC Corporate Bankingate(lines):
        if 'Balance Summary' in line:
            # Next lines contain beginning/ending balances     
    return result
ate(lines):
        if 'Balance Summary' in line:
            # Next lines contain beginning/ending balances
            for j in range(i + 1, min(i + 5, len(lines))):
                if 'Beginning' in lines[j]:
                    match = re.search(r'([\d,]+\.?\d*)', lines[j])
                    if match:
                        result['beginning_balance'] = float(
                            match.group(1).replace(',', '')
                        )
                if 'Ending' in lines[j]:
                    match = re.search(r'([\d,]+\.?\d*)', lines[j])
                    if match:
                        result['ending_balance'] = float(
                            match.group(1).replace(',', '')
                        )

    # Extract deposits
    _extract_pnc_deposits(text, result)

    # Extract checks
    _extract_pnc_checks(text, result)

    # Extract ACH debits
    _extract_pnc_ach_debits(text, result)

    # Extract ledger balances
    _extract_pnc_ledger_balances(text, result)


def _extract_pnc_deposits(text: str, result: Dict[str, Any]) -> None:
    """Extract deposits from PNC statement."""
    lines = text.split('\n')
    in_deposits = False

    for i, line in enumerate(lines):
        if 'Deposits 1 transaction' in line or 'Deposits and Other Credits' in line:
            in_deposits = True
            continue

        if in_deposits:
            if 'posted' in line and 'Amount' in line:
                # Header line, next lines are transactions
                for j in range(i + 1, min(i + 10, len(lines))):
                    match = re.search(
                        r'(\d{2}/\d{2})\s+([\d,]+\.?\d*)\s+(.+?)\s+(\d+)',
                        lines[j],
                    )
                    if match:
                        deposit = {
                            'date': match.group(1),
                            'amount': float(match.group(2).replace(',', '')),
                            'description': match.group(3).strip(),
                            'reference': match.group(4),
                        }
                        result['deposits'].append(deposit)
                        result['transactions'].append(
                            {
                                'type': 'deposit',
                                'date': match.group(1),
                                'amount': float(match.group(2).replace(',', '')),
                                'description': match.group(3).strip(),
                            }
                        )

            if 'Funds Transfer' in line or 'Checks and Other' in line:
                break


def _extract_pnc_checks(text: str, result: Dict[str, Any]) -> None:
    """Extract checks from PNC statement.

    PNC uses a 3-column grid layout for checks:
      date check_num amount ref  date check_num amount ref  date check_num amount ref
    """
    lines = text.split('\n')
    in_checks = False

    for i, line in enumerate(lines):
        if 'Checks and Substitute Checks' in line:
            in_checks = True
            continue

        if in_checks:
            # Skip header lines
            if 'posted' in line.lower() or 'Date' in line and 'Check' in line:
                continue

            # Find all check entries in the line (3-column grid)
            # Pattern: mm/dd check_num amount reference_num
            matches = re.findall(
                r'(\d{2}/\d{2})\s+(\d{3,5})\s+([\d,]+\.\d{2})\s+(\d+)',
                line,
            )
            for m in matches:
                check = {
                    'date': m[0],
                    'check_number': m[1],
                    'amount': float(m[2].replace(',', '')),
                    'reference': m[3],
                }
                result['checks'].append(check)
                result['withdrawals'].append(check)
                result['transactions'].append(
                    {
                        'type': 'check',
                        'date': m[0],
                        'amount': -float(m[2].replace(',', '')),
                        'check_number': m[1],
                        'description': f'Check #{m[1]}',
                    }
                )

            if 'ACH Debits' in line or 'Corporate ACH' in line:
                break
atch.group(1),
                    'amount': float(match.group(2).replace(',', '')),
                    'description': desc,
                    'reference': match.group(4),
                }
                result['ach_debits'].append(ach)
                result['withdrawals'].append(ach)
                result['transactions'].append({
                    'type': 'ach_debit',
                    'date': match.group(1),
                    'amount': -float(match.group(2).replace(',', '')),
                    'description': desc,
                })

            if 'Member FDIC' in line or ('Ending balance' in line):
                break

lance'] = float(
                    match.group(1).replace(',', '')
                )
                break

    # Bank of America statement in sample shows no transactions
    result['transactions'] = []


def _parse_keybank(text: str, result: Dict[str, Any]) -> None:
    """
    Parse KeyBank statement.

    Args:
        text: Extracted text from PDF page
        result: Result dictionary to populate
    """
    lines = text.split('\n')

    # Extract account number
    for line in lines:
        if 'Commercial Control Transaction' in line:
            match = re.search(r'(\d+)', line.split()[-1])
            if match:
                result['account_number'] = match.group(1)
                break

    - Skip header lines
    for line in lines:
        if 'for' in line.lower() and 'to' in line.lower():
            match = re.search(h
                r&(\x + \d+, \d{4})\s+to\s+(\w+ \d+, \d{4})',
                line,
            )
            if match:
                result['statement_period'] = {
                    'start': match.group(1),
                    'end': match.group(2),
                }
                break

   # Extract beginning balance
    for line in lines:
        if 'Beginning balance' in line.lower():
            match = re.search(r'\$([\d,]+\.?\d*)', line)
            if match:
                result['beginning_balance'] = float(
                    match.group(1).replace(',', '')
                )
                break

    # Extract ending balance
    for line in lines:
        if 'Ending balance' in line.lower():
            match = re.search(r \$([\d,]+\.?\d*)', line)
            if match:
                result['ending_balance'] = float(
                    match.group(1).replace(',', '')
                )
                break

    # Extract deposits
    _extract_keybank_deposits(text, result)

    # Extract withdrawals
    _extract_keybank_withdrawals(text, result)

    # Extract fees
    _extract_keybank_fees(text, result)

ot appear to be a recognized bank statement"
                )

            # Check for account information
            if 'Account' not in text and 'account' not in text:
                issues.append("Missing account information")

    except Exception as e:
        return False, [f"Failed to open PDF file: {str(e)}"]

    return len(issues) == 0, issues

