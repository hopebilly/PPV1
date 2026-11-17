const LocalStorageVersion: string = "v2"; //Only change to force end user's cache to be refreshed.
export class LocalStorageKeys {
  static readonly Employees: string = `NIHNCIFRED-EMPLOYEES-${LocalStorageVersion}`;
}

export class LocalStorageHelper {
  static getKeyString_(key: string, siteUrl: string): string {
    const keyString = `${key}-${siteUrl}`;
    return keyString;
  }

  public static set<T = object>(key: string, siteUrl: string, storeState: T): boolean {
    if (!localStorage) {
      return false;
    }

    try {
      const serializedState = JSON.stringify(storeState);
      localStorage.setItem(this.getKeyString_(key, siteUrl), serializedState);
      window.dispatchEvent(new Event("storage"));
      return true;
    } catch (error) {
      throw new Error("store serialization failed");
    }
  }

  public static remove(key: string, siteUrl: string): boolean {
    if (!localStorage) {
      return false;
    }

    try {
      localStorage.removeItem(this.getKeyString_(key, siteUrl));
      window.dispatchEvent(new Event("storage"));
      return true;
    } catch (error) {
      throw new Error("store item removal failed");
    }
  }

  public static get<T = object>(key: string, siteUrl: string): T | undefined {
    if (!localStorage) {
      return;
    }

    try {
      const serializedState = localStorage.getItem(this.getKeyString_(key, siteUrl));
      if (serializedState === null) {
        return;
      }
      return JSON.parse(serializedState);
    } catch (error) {
      throw new Error("store deserialization failed");
    }
  }
}
