import os
from azure.storage.blob import BlockBlobService
import pathlib

class _blob_client:
    def __init__(self):
        # check if the home directory path is valid
        self.path = "C:\\Users\\mech-user\\Desktop\\IoT\\pi"
        self.p=pathlib.Path('/pi/')
        if not os.path.isdir("C:\\Users\\mech-user\\Desktop\\IoT\\pi"):
            raise Exception("home directory is not found")

        # build client
        self.client = BlockBlobService(
            account_name="shuto1441", account_key="4spqDCxOUGdpNyYP6NNM1q+RkX9NYvQgB0UjIynPpt/nbs9mB8Q65iz3wSYcMldkyOhj5lbQQHCPGBAlVT3rxg==")
        """
        # set proxy if enabled
        if conf["proxy_enabled"]:
            self.client.set_proxy(conf["proxy_addr"], conf["proxy_port"],
                                  user=conf["proxy_user"],
                                  password=conf["proxy_pass"])
        """

        # set container name with checking
        if "quickstart" in [x.name for x in self.client.list_containers()]:
            self.__container = "quickstart"
        else:
            raise Exception("invalid container name")

    def download(self,blob):
        """
        # download all blobs to local file
        for blob in self.fetch_remote():
        """
        file_path = blob
        # download blob
        print("receiving blob: {}".format(blob))
        self.client.get_blob_to_path(container_name=self.__container,
                                        blob_name=blob,
                                        file_path=file_path)
    def upload(self):
        # upload all local file to remote blob
        for item in os.listdir(self.path):
            local_path = os.path.join(self.path, item)
            if os.path.isfile(local_path):
                # upload file to blob
                print("uploading file: {}".format(item))
                self.client.create_blob_from_path(container_name=self.__container,
                                                  blob_name=item,
                                                  file_path=local_path)
    def fetch_remote(self):
        # list all blob name in the container
        print("getting blob list...")
        return [x.name for x in self.client.list_blobs(self.__container)]
    
    def fetch_local(self):
        file_info = []
        for root, dirs, files in os.walk(self.path):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                file_path=str(file_path)
                file_path=file_path.split("pi")[-1]
                file_info.append(file_path)
 
        return file_info

    def clear(self,blob):
        """
        # remove all blobs in the container
        for blob in self.fetch_remote():
        """
        print("removing blob: {}".format(blob))
        self.client.delete_blob(container_name=self.__container,
                                blob_name=blob)

"""
client = _blob_client()
local_ = client.fetch_local()
remote_ = client.fetch_remote()
download_blobs = []
remote_blobs=[]
remove_blobs=[]


# debug print
print("local ----")
for item in local_:
    print(item)
print("remote ---")
for item in remote_:
    print(item)


#Azureの整理
for item_remote in remote_:
    remote_blobs.append(item_remote)
remote_blobs.sort()
for i in range(len(remote_blobs)-1,1,-1):
    if remote_blobs[i][0:3]==remote_blobs[i-1][0:3]:
        remove_blobs.append(remote_blobs[i-1])
for blob in remove_blobs:
    print("remove: {}".format(blob))
    client.clear(blob)

#Azureからのダウンロード
for item_remote in remote_:
    for item_local in local_:
        if str(item_local)[1:4]==str(item_remote)[3:6] and str(item_local)!=str(item_remote):
             download_blobs.append(item_remote)
             os.remove("C:\\Users\\mech-user\\Desktop\\IoT\\ダウンロードデータ\\"+item_local)
for blob in download_blobs:
    print("download: {}".format(blob))
    client.download(blob)
"""