from typing import List, Dict, Union
import logging
from http_client import HttpClient
from exceptions import HermesMSGraphError

logger = logging.getLogger(__name__)

class PlannerService:
    def __init__(self, http_client: HttpClient):
        self.http = http_client

    def _process_plans(self, plans: List[Dict]) -> List[Dict[str, str]]:
        """Helper method to process plan data."""
        return [
            {
                "id": plan["id"],
                "title": plan["title"],
                "owner": plan.get("owner", {}).get("user", {}).get("displayName", "N/A")
            }
            for plan in plans
        ]

    def _fetch_data(self, url: str) -> List[Dict]:
        """Helper method to fetch data from a given URL with error handling."""
        response = self.http.get(url)
        if response.status_code != 200:
            raise HermesMSGraphError(f"Failed to fetch data: {response.status_code} - {response.text}")
        return response.json().get("value", [])

    def list_plans_by_group_id(self, group_id: str, data: str = "all") -> Union[List[Dict], List[Dict[str, str]]]:
        """
        List plans by group ID.
        :param group_id: The ID of the group.
        :param data: "all" to return raw data, or "processed" to return simplified data.
        :return: List of plans.
        """
        url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/planner/plans"
        plans = self._fetch_data(url)
        return plans if data == "all" else self._process_plans(plans)

    def list_visible_plans_by_user_id(self, user_id: str) -> List[Dict[str, str]]:
        """
        List visible plans by user ID.
        :param user_id: The ID of the user.
        :return: List of processed plans.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/planner/plans"
        plans = self._fetch_data(url)
        return self._process_plans(plans)

    def list_tasks_by_user_id(self, user_id: str) -> List[Dict]:
        """
        List tasks assigned to a user by their user ID.
        :param user_id: The ID of the user.
        :return: List of tasks.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/planner/tasks"
        tasks = self._fetch_data(url)
        logger.debug(f"Fetched tasks for user {user_id}: {tasks}")
        return tasks