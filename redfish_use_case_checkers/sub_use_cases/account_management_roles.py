# Copyright Notice:
# Copyright 2017-2025 Distributed Management Task Force, Inc. All rights reserved.
# License: BSD 3-Clause License. For full text see link: https://github.com/DMTF/Redfish-Use-Case-Checkers/blob/main/LICENSE.md

"""Account Management role sub-use cases."""

import redfish_utilities

from redfish_use_case_checkers import logger
from redfish_use_case_checkers.system_under_test import SystemUnderTest


EXPECTED_PRIVILEGES = {
    "Administrator": {
        "Login",
        "ConfigureManager",
        "ConfigureUsers",
        "ConfigureSelf",
        "ConfigureComponents",
    },
    "Operator": {"Login", "ConfigureSelf", "ConfigureComponents"},
    "ReadOnly": {"Login", "ConfigureSelf"},
}


def check_assigned_privileges(sut: SystemUnderTest, category_name: str, test_name: str):
    """Verifies the assigned privileges of the standard predefined roles."""

    operation = "Locating the role collection"
    logger.logger.info(operation)
    try:
        account_service_uri = sut.service_root["AccountService"]["@odata.id"]
        account_service = sut.session.get(account_service_uri)
        redfish_utilities.verify_response(account_service)
        role_collection_uri = account_service.dict["Roles"]["@odata.id"]
        role_collection = sut.session.get(role_collection_uri)
        redfish_utilities.verify_response(role_collection)
    except Exception as err:
        sut.add_test_result(
            category_name,
            test_name,
            operation,
            "FAIL",
            "Failed to get the role collection ({}).".format(err),
        )
        return

    roles = {}
    for member in role_collection.dict.get("Members", []):
        if not isinstance(member, dict) or "@odata.id" not in member:
            continue
        try:
            role = sut.session.get(member["@odata.id"])
            redfish_utilities.verify_response(role)
            role_id = role.dict.get("RoleId")
            if role_id:
                roles[role_id] = role.dict
        except Exception as err:
            operation = "Getting role '{}'".format(member["@odata.id"])
            sut.add_test_result(
                category_name,
                test_name,
                operation,
                "FAIL",
                "Failed to get role ({}).".format(err),
            )

    for role_name, expected_privileges in EXPECTED_PRIVILEGES.items():
        operation = "Checking assigned privileges for '{}'".format(role_name)
        logger.logger.info(operation)
        if role_name not in roles:
            sut.add_test_result(
                category_name,
                test_name,
                operation,
                "FAIL",
                "Role '{}' is not present in the role collection.".format(role_name),
            )
            continue

        assigned_privileges = roles[role_name].get("AssignedPrivileges")
        if not isinstance(assigned_privileges, list):
            sut.add_test_result(
                category_name,
                test_name,
                operation,
                "FAIL",
                "Role '{}' does not contain an AssignedPrivileges array.".format(role_name),
            )
            continue

        actual_privileges = set(assigned_privileges)
        if actual_privileges == expected_privileges and len(assigned_privileges) == len(actual_privileges):
            sut.add_test_result(category_name, test_name, operation, "PASS")
        else:
            sut.add_test_result(
                category_name,
                test_name,
                operation,
                "FAIL",
                "Role '{}' has AssignedPrivileges {}; expected {}.".format(
                    role_name,
                    sorted(actual_privileges),
                    sorted(expected_privileges),
                ),
            )