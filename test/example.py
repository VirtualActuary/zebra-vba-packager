import locate
locate.allow_relative_location_imports("..")
import zebra_vba_packager
import zebra_vba_packager.zebra_config
from zebra_vba_packager.zebra_config import Source, Config


def pre_process(source):
    print("\n*** Pre process: ***", source.url_source)

def mid_process(source):
    print("\n*** Mid process: ***")
    for i in source.temp_transformed.rglob("*"):
        print(i)

def post_process(source):
    print("\n*** Post Process: ***")
    print("We are done!")


Config(
    Source(
        pre_process=pre_process,

        url_source="https://github.com/sdkn104/VBA-CSV/archive/refs/tags/v1.9.zip",
        glob_include=['**/*.bas', '**/*.cls'],
        glob_exclude=['**/*Example.bas', '**/*Test.bas'],

        mid_process=mid_process,
        post_process=post_process
    ),
    Source(
        git_source="https://github.com/ws-garcia/VBA-CSV-interface.git",
        git_rev="795e583701efece8d14d2f16693ad1673fabe042",
        glob_include=['**/*.bas', '**/*.cls'],
        glob_exclude=['**/Tests/'],
        mid_process=mid_process
    )
).run()
