#!/usr/bin/env python3
"""Filter junk 'emails' picked up by the website-enrichment regex.

The generic email regex also matches error-tracking IDs and image-file
names embedded in page HTML, e.g.:
  aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa@sentry-next.wixpress.com
  aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa@sentry.io
  boels_logo_web@2x.png / logo@2x.png / media--...--query@2x.jpg
  sdk@0.1.26-site.compat.min.js
None of these can receive mail. Drop them wherever an enriched
邮箱 / 全部邮箱 / 主邮箱 value is built or cleaned.
"""
import re

_JUNK_TLDS = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'js', 'css', 'map'}
_HEX32 = re.compile(r'^[0-9a-f]{32}$')


def is_junk_email(addr):
    """True if addr is a tracking ID / file name, not a real mailbox."""
    if not addr or '@' not in addr:
        return True
    addr = addr.strip()
    local, dom = addr.rsplit('@', 1)
    local, dom = local.strip(), dom.strip().lower()
    if not local or not dom or '.' not in dom:
        return True
    if 'sentry' in dom:
        return True
    tld = dom.rsplit('.', 1)[-1]
    if tld in _JUNK_TLDS:
        return True
    if dom.endswith('.de.de'):
        return True
    if _HEX32.match(local):
        return True
    return False


def clean_email_list(value):
    """Remove junk entries from a '; '-joined email string."""
    if not value:
        return ''
    parts = [p.strip() for p in str(value).split(';') if p.strip()]
    return '; '.join(p for p in parts if not is_junk_email(p))


# UK contact-person fragments: the title-first regex often captures the
# words after a title instead of a name, e.g. 'Close Project Project',
# 'Kieran Carville General', 'Shaun Lawn'. Keep only clean 2-3 word names.
_PERSON_BAD_WORDS = {
    'director', 'manager', 'general', 'project', 'services', 'financial',
    'operations', 'depot', 'team', 'office', 'sales', 'lawn', 'colchester',
    'jnr', 'snr', 'ltd', 'limited', 'company', 'group', 'contact', 'email',
}
_PERSON_RE = re.compile(r"^[A-Z][a-z]+(?: [A-Z][a-z]+){1,2}$")


def clean_uk_person(name):
    """Return the name if it looks like a real 2-3 word personal name, else ''."""
    if not name:
        return ''
    name = ' '.join(name.strip().split())
    if not _PERSON_RE.match(name):
        return ''
    if any(w.lower() in _PERSON_BAD_WORDS for w in name.split()):
        return ''
    return name
