extern crate embed_resource;

fn main() {
    embed_resource::compile("schedule-sort-manifest.rc", embed_resource::NONE)
        .manifest_optional()
        .unwrap();
}
