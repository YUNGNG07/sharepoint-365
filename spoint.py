from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.activities.entity import SPActivityEntity
from office365.sharepoint.client_context import ClientContext
import time
import json

class Spoint():
    def __init__(self):
        self.__spoint_init()

    def __spoint_init(self):
        """
        Initialize sharepoint context.
        """
        client_id = '3268db4f-6258-4fae-bc12-7982c1eb1783'
        client_secret = 'HYtKhZwvJ03862cL16MiAOr5V41LAfaZDPHMTqmJh3I='
        site_url = 'https://keysighttech.sharepoint.com/sites/TestProjects-InventoryManagementSystem'

        while True:
            try:
                credentials = ClientCredential(client_id, client_secret)
                self.context = ClientContext(site_url).with_credentials(credentials)
                target_web = self.context.web.get().execute_query()
                break
            except Exception as e:
                print('[SPOINT_INIT] init with error: ' + str(e.__class__) + ' retrying')
                time.sleep(1)
        print(f'[SPOINT_INIT] target url: {target_web.url}')

    def get_items_title(self):
        LIST_TITLE = 'Inventory Management System'
        source_list = self.context.web.lists.get_by_title(LIST_TITLE)
        items = source_list.items.get().execute_query()

        for item in items:
            print(item.properties.get('Title'))

    def list_site_pages(self):
        site_pages = self.context.site_pages.pages.get().execute_query()
        for site_page in site_pages:
            print(site_page.file_name)

    def get_web_activity(self):
        activities = self.context.web.activities.get().execute_query()
        for activity in activities:
            print(activity.action.facet_type)

    def get_all_activity(self):
        webs = self.context.web.get_all_webs().execute_query()
        for web in webs:
            print(web.url)

    def get_roles(self):
        role_defs = self.context.web.role_definitions.get().execute_query()
        for role_def in role_defs:
            print(role_def.name)

    def read_custom_items(self):
        view = self.context.web.default_document_library().views.get_by_title('All Documents')
        items = view.get_items().expand(['Author']).execute_query()
        for item in items:
            print(item.properties)

    def export_site_users(self):
        users = self.context.web.site_users.select(['LoginName']).get().top(100).execute_query()
        for user in users:
            print(user.login_name)

    def whoami(self):
        whoami = self.context.web.current_user.get().execute_query()
        print(json.dumps(whoami.to_json(), indent=4))

    def get_from_list(self):
        target_list = self.context.web.default_document_library()
        result = target_list.get_site_script().execute_query()
        print(result.value)

    def get_admins(self):
        result = self.context.site.get_site_administrators().execute_query()
        for info in result.value:
            print(info)

    def get_list_permissions(self):
        doc_lib = self.context.web.default_document_library()
        result = doc_lib.get_user_effective_permissions(self.context.web.current_user).execute_query()
        print(result.value.permission_levels)

    def get_page_content(self):
        file = self.context.web.get_file_by_server_relative_path('SitePages/Home.aspx')
        file_item = (file.listItemAllFields.select(['CanvasContent1', 'LayoutWebpartsContent']).get().execute_query())
        print(file_item.properties.get('CanvasContent1'))

    def export_top_navigation(self):
        nav = self.context.web.navigation.top_navigation_bar.get().execute_query()
        print(json.dumps(nav.to_json(), indent=4))

if __name__ == '__main__':
    try:
        spoint = Spoint()
        # spoint.get_items_title()
        # spoint.list_site_pages()
        # spoint.get_web_activity()
        # spoint.get_all_activity()
        # spoint.get_roles()
        # spoint.read_custom_items()
        # spoint.export_site_users()
        # spoint.whoami()
        # spoint.get_from_list()
        # spoint.get_admins()
        # spoint.get_list_permissions()
        # spoint.get_page_content()
        spoint.export_top_navigation()
    except Exception as e:
        print(f'Error: {str(e)}')
