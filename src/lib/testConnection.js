import { supabase } from "./supabase";

export async function testConnection() {
  const { data, error } = await supabase
    .from("items")
    .select("*")
    .limit(1);

  console.log("DATA:", data);
  console.log("ERROR:", error);
}
